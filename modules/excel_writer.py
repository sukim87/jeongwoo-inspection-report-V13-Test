from io import BytesIO
from pathlib import Path
import math
import re
import tempfile
from copy import copy

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment, Font
from PIL import Image, ImageOps, ImageFilter, ImageEnhance

TEMPLATE = Path(__file__).resolve().parents[1] / 'templates' / '검사결과서.xlsx'
SLOTS = [
    ('검사결과서', 'A21'), ('검사결과서', 'G21'),
    ('검사결과서(을지)', 'A11'), ('검사결과서(을지)', 'G11'),
    ('검사결과서(을지)', 'A24'), ('검사결과서(을지)', 'G24'),
    ('검사결과서(을지)', 'A37'), ('검사결과서(을지)', 'G37'),
]
# Exact display size for the merged photo frames in the workbook.
PHOTO_WIDTH = 360
PHOTO_HEIGHT = 340
CANVAS = (1413, 1200)  # same aspect as the Excel photo frame


def safe(value):
    return re.sub(r'[\\/:*?"<>|]+', '-', str(value or '미입력')).strip()


def _full_photo_canvas(image):
    """Keep the complete photo while filling the frame with a soft background."""
    background = ImageOps.fit(image, CANVAS, Image.Resampling.LANCZOS)
    background = background.filter(ImageFilter.GaussianBlur(radius=24))
    background = ImageEnhance.Brightness(background).enhance(0.74)
    foreground = ImageOps.contain(image, CANVAS, Image.Resampling.LANCZOS)
    left = (CANVAS[0] - foreground.width) // 2
    top = (CANVAS[1] - foreground.height) // 2
    background.paste(foreground, (left, top))
    return background


def photo_canvas(uploaded, degree, mode):
    uploaded.seek(0)
    image = ImageOps.exif_transpose(Image.open(uploaded)).convert('RGB')
    if degree:
        image = image.rotate(-degree, expand=True)
    if mode == '사진 칸 채우기':
        return ImageOps.fit(image, CANVAS, Image.Resampling.LANCZOS, centering=(0.5, 0.5))
    return _full_photo_canvas(image)


def _fit_text(cell, text, base_size=9, min_size=6.5, chars_per_line=52, max_lines=5):
    """Wrap long inspection text and reduce font only when necessary."""
    text = str(text or '')
    cell.value = text
    estimated_lines = max(1, math.ceil(len(text) / chars_per_line) + text.count('\n'))
    font_size = max(min_size, base_size - max(0, estimated_lines - max_lines) * 0.55)
    font = copy(cell.font)
    font.size = font_size
    cell.font = font
    cell.alignment = Alignment(
        horizontal='left', vertical='center', wrap_text=True,
        shrink_to_fit=True
    )
    return estimated_lines


def build_report(data, photos, rotations, modes=None):
    modes = modes or ['전체 사진 표시'] * len(photos)
    wb = load_workbook(TEMPLATE)
    quick = wb['간편입력']
    main = wb['검사결과서']
    extra = wb['검사결과서(을지)']

    # Open the actual inspection report first and keep the data-entry sheet hidden.
    quick.sheet_state = 'hidden'
    for report_sheet in (main, extra):
        # 0.60 inch is about 15 mm, leaving room for binder punching.
        report_sheet.page_margins.top = 0.60
        report_sheet.page_setup.orientation = 'portrait'
        report_sheet.page_setup.fitToWidth = 1
        report_sheet.page_setup.fitToHeight = 1
        report_sheet.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = wb.sheetnames.index('검사결과서')
    main.sheet_view.selection[0].activeCell = 'A1'
    main.sheet_view.selection[0].sqref = 'A1'

    quick_values = {
        'B5': data['customer'], 'B6': data['part_name'], 'B7': data['order_no'],
        'B8': data['po_step_no'], 'B9': data['date'], 'B10': data['types'],
        'B11': data['spec'], 'B14': data['inspection_content'], 'B15': data['judgment'],
        'B17': data['result'], 'B18': data['action'], 'B19': data['caption'],
        'B20': data['inspector'], 'B21': data['inspection_location'],
    }
    for cell, value in quick_values.items():
        quick[cell] = value

    main_values = {
        'J3': data['doc'], 'C5': data['customer'], 'I5': data['types'],
        'C6': data['part_name'], 'I6': data['po_step_no'], 'C7': data['order_no'],
        'I7': data['date'], 'A11': data['spec'], 'J11': data['inspection_location'],
        'A33': data['caption'], 'A37': data['inspector'], 'J37': data['judgment'],
    }
    for cell, value in main_values.items():
        main[cell] = value

    # Prevent long inspection criteria/content/result/action from being clipped.
    content_lines = _fit_text(main['F11'], data['inspection_content'], base_size=8.5, min_size=6.5, chars_per_line=27, max_lines=4)
    _fit_text(main['A11'], data['spec'], base_size=8.5, min_size=6.5, chars_per_line=34, max_lines=4)
    _fit_text(main['J11'], data['inspection_location'], base_size=8.5, min_size=7, chars_per_line=15, max_lines=3)
    _fit_text(main['C13'], data['result'], base_size=8.5, min_size=6.3, chars_per_line=65, max_lines=6)
    _fit_text(main['C16'], data['action'], base_size=8.5, min_size=6.5, chars_per_line=70, max_lines=4)
    main.row_dimensions[11].height = min(78, max(38, 16 * min(content_lines, 5)))

    extra_values = {
        'J3': data['doc'], 'C5': data['customer'], 'I5': data['types'],
        'C6': data['part_name'], 'I6': data['po_step_no'], 'C7': data['order_no'],
        'I7': data['date'],
    }
    for cell, value in extra_values.items():
        extra[cell] = value

    for ws in (main, extra):
        ws._images = []

    with tempfile.TemporaryDirectory() as temp_dir:
        for index, uploaded in enumerate(photos[:8]):
            degree = rotations[index] if index < len(rotations) else 0
            mode = modes[index] if index < len(modes) else '전체 사진 표시'
            image = photo_canvas(uploaded, degree, mode)
            path = f'{temp_dir}/{index}.jpg'
            image.save(path, 'JPEG', quality=93, optimize=True)
            excel_image = XLImage(path)
            excel_image.width = PHOTO_WIDTH
            excel_image.height = PHOTO_HEIGHT
            sheet_name, anchor = SLOTS[index]
            row, col = coordinate_to_tuple(anchor)
            excel_image.anchor = OneCellAnchor(
                _from=AnchorMarker(col=col-1, colOff=pixels_to_EMU(18), row=row-1, rowOff=pixels_to_EMU(8)),
                ext=XDRPositiveSize2D(pixels_to_EMU(PHOTO_WIDTH), pixels_to_EMU(PHOTO_HEIGHT)),
            )
            wb[sheet_name].add_image(excel_image)

        output = BytesIO()
        wb.save(output)

    return output.getvalue(), f"{safe(data['doc'])}.xlsx"
