from io import BytesIO
from pathlib import Path
import re
from openpyxl import load_workbook
from PIL import Image, ImageOps, ImageEnhance, ImageFilter

ORDER_RE = re.compile(r'\bH\d{6}[A-Z]{2}\b', re.I)
PO_RE = re.compile(r'\bPS\d{7}(?:\s*[,/]\s*PS\d{7})*\b', re.I)
BAD_CHARS = re.compile(r'[^0-9A-Za-z가-힣㈜()\[\].,&/+\- ]')


def _after(ws, label, max_columns=32):
    for row in ws.iter_rows():
        for cell in row:
            if label in str(cell.value or ''):
                # Prefer values to the right of the label, skipping blank merged cells.
                for col in range(cell.column + 1, min(ws.max_column, cell.column + max_columns) + 1):
                    value = ws.cell(cell.row, col).value
                    if value not in (None, ''):
                        text = str(value).strip()
                        if label not in text:
                            return text
    return ''


def _valid_name(text):
    text = re.sub(r'\s+', ' ', str(text or '')).strip(' |:_-.,')
    if len(text) < 2 or len(text) > 80:
        return ''
    if BAD_CHARS.search(text):
        return ''
    alnum = sum(ch.isalnum() or ('가' <= ch <= '힣') for ch in text)
    if alnum / max(len(text), 1) < 0.55:
        return ''
    # Reject OCR fragments dominated by one-character tokens.
    tokens = text.split()
    if len(tokens) >= 5 and sum(len(t) <= 2 for t in tokens) / len(tokens) > 0.55:
        return ''
    return text


def _normalize_order(text):
    match = ORDER_RE.search(str(text or '').upper().replace(' ', ''))
    return match.group(0).upper() if match else ''


def _normalize_po(text):
    compact = str(text or '').upper().replace(' ', '')
    matches = re.findall(r'PS\d{7}', compact)
    return ','.join(dict.fromkeys(matches))


def read_xlsx(raw):
    ws = load_workbook(BytesIO(raw), data_only=True).worksheets[0]
    info = {
        'customer': _valid_name(_after(ws, '고객사')),
        'part_name': _valid_name(_after(ws, '품목명') or _after(ws, '부품명')),
        'order_no': _normalize_order(_after(ws, '수주번호')),
        'po_step_no': _normalize_po(_after(ws, '발주/공정번호')),
        'inspection_location': _valid_name(_after(ws, '업체/팀명')),
        'review_fields': [],
    }
    info['review_fields'] = [key for key in ('customer','part_name','order_no','po_step_no','inspection_location') if not info[key]]
    return info, None


def _render(raw):
    import fitz
    doc = fitz.open(stream=raw, filetype='pdf')
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
    return Image.open(BytesIO(pix.tobytes('png'))).convert('RGB'), page.get_text('text')


def _ocr_full(image):
    import pytesseract
    gray = ImageOps.grayscale(image)
    gray = ImageEnhance.Contrast(gray).enhance(1.8).filter(ImageFilter.SHARPEN)
    try:
        return pytesseract.image_to_string(gray, lang='kor+eng', config='--oem 3 --psm 6')
    except Exception:
        return pytesseract.image_to_string(gray, lang='eng', config='--oem 3 --psm 6')


def _value_after_label(text, labels):
    lines = [re.sub(r'\s+', ' ', line).strip() for line in str(text or '').splitlines() if line.strip()]
    for i, line in enumerate(lines):
        for label in labels:
            if label.lower() in line.lower():
                tail = re.split(re.escape(label), line, flags=re.I, maxsplit=1)[-1].strip(' :|')
                if tail and tail.lower() not in ('customer','item name','supplier/team','order no.','po/step no.'):
                    return tail
                for next_line in lines[i+1:i+4]:
                    if not any(x.lower() in next_line.lower() for x in ('prepared by','approved by','작성','승인')):
                        return next_line
    return ''


def _parse_text(text):
    order = _normalize_order(text)
    po = _normalize_po(text)
    customer = _valid_name(_value_after_label(text, ['고객사', 'Customer']))
    part = _valid_name(_value_after_label(text, ['품목명', '부품명', 'Item Name', 'Part Name']))
    location = _valid_name(_value_after_label(text, ['업체/팀명', 'Supplier/Team']))
    # Do not let field headers themselves become values.
    forbidden = {'customer','item name','part name','supplier/team','order no.','po/step no.','고객사','품목명','부품명','업체/팀명'}
    customer = '' if customer.lower() in forbidden else customer
    part = '' if part.lower() in forbidden else part
    location = '' if location.lower() in forbidden else location
    return {'customer':customer,'part_name':part,'order_no':order,'po_step_no':po,'inspection_location':location}


def read_pdf(raw):
    preview, embedded_text = _render(raw)
    embedded = _parse_text(embedded_text)
    # OCR is only a fallback. It never overrides a verified embedded-text value.
    if not embedded['order_no'] or not embedded['po_step_no'] or not embedded['inspection_location']:
        ocr = _parse_text(_ocr_full(preview))
        for key in embedded:
            if not embedded[key]:
                embedded[key] = ocr[key]
    embedded['order_no'] = _normalize_order(embedded['order_no'])
    embedded['po_step_no'] = _normalize_po(embedded['po_step_no'])
    for key in ('customer','part_name','inspection_location'):
        embedded[key] = _valid_name(embedded[key])
    embedded['review_fields'] = [key for key in ('customer','part_name','order_no','po_step_no','inspection_location') if not embedded[key]]
    return embedded, preview


def read_application(uploaded):
    raw = uploaded.getvalue()
    return read_xlsx(raw) if Path(uploaded.name).suffix.lower() == '.xlsx' else read_pdf(raw)
