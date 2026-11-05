from email.message import EmailMessage
from html import escape
import re


def split_addresses(value):
    return [x.strip() for x in re.split(r'[;,\n]+', value or '') if x.strip()]


def judgment_mark(judgment):
    return {
        '합격': '✅ 합격',
        '불합격': '❌ 불합격',
        '조건부 합격': '⚠️ 조건부 합격',
        'N/A': 'N/A',
    }.get(judgment, judgment or '')


def inspector_name(data):
    return str(data.get('inspector') or '검사자').strip()


def make_subject(data):
    return (
        f"[검사결과 공유] {data.get('order_no','')} / "
        f"{data.get('po_step_no','')} / {data.get('judgment','')} / "
        f"{data.get('document_no','')}"
    )


def make_html_body(data):
    # IMPORTANT: return compact HTML without leading indentation/newlines.
    # This prevents Streamlit Markdown from displaying the HTML as a code block.
    e = escape
    inspector = e(inspector_name(data))
    order_no = e(str(data.get('order_no', '')))
    part_name = e(str(data.get('part_name', '')))
    inspection_types = e(str(data.get('types', '')))
    location = e(str(data.get('inspection_location', '')))
    judgment = e(judgment_mark(data.get('judgment', '')))
    inspection_content = e(str(data.get('inspection_content', '')))

    return (
        '<div style="font-family:Malgun Gothic,Arial,sans-serif;font-size:11pt;line-height:1.7;color:#222">'
        f'<p>안녕하십니까.<br>품질팀 <b>{inspector}</b>입니다.</p>'
        '<p>제목의 건 검사결과를 아래와 같이 공유드립니다.</p>'
        '<div style="border-left:5px solid #c8102e;background:#f6f7f9;padding:15px 18px;margin:16px 0;max-width:760px">'
        '<div style="font-size:13pt;font-weight:700;margin-bottom:10px">검사결과 요약</div>'
        f'<div><b>수주번호:</b> {order_no}</div>'
        f'<div><b>품명:</b> {part_name}</div>'
        f'<div><b>검사종류:</b> {inspection_types}</div>'
        f'<div><b>검사장소:</b> {location}</div>'
        f'<div><b>최종판정:</b> <b>{judgment}</b></div>'
        '</div>'
        '<div style="margin:16px 0;max-width:760px">'
        '<div style="font-size:12pt;font-weight:700;margin-bottom:6px">검사내용</div>'
        f'<div>{inspection_content}</div>'
        '</div>'
        '<p><b>상세 검사내용 및 전체 검사사진은 첨부된 검사결과서를 참고하여 주시기 바랍니다.</b></p>'
        f'<p>감사합니다.<br><b>{inspector}</b> 드림.</p>'
        '</div>'
    )


def make_text_body(data):
    inspector = inspector_name(data)
    return (
        f"안녕하십니까.\n품질팀 {inspector}입니다.\n\n"
        "제목의 건 검사결과를 아래와 같이 공유드립니다.\n\n"
        "검사결과 요약\n"
        f"수주번호: {data.get('order_no','')}\n"
        f"품명: {data.get('part_name','')}\n"
        f"검사종류: {data.get('types','')}\n"
        f"검사장소: {data.get('inspection_location','')}\n"
        f"최종판정: {judgment_mark(data.get('judgment',''))}\n\n"
        "검사내용\n"
        f"{data.get('inspection_content','')}\n\n"
        "상세 검사내용 및 전체 검사사진은 첨부된 검사결과서를 참고하여 주시기 바랍니다.\n\n"
        f"감사합니다.\n{inspector} 드림.\n"
    )


def build_eml(data, excel_bytes, excel_filename, to_value='', cc_value=''):
    message = EmailMessage()
    message['Subject'] = make_subject(data)
    to_list = split_addresses(to_value)
    cc_list = split_addresses(cc_value)
    if to_list:
        message['To'] = ', '.join(to_list)
    if cc_list:
        message['Cc'] = ', '.join(cc_list)
    message.set_content(make_text_body(data))
    message.add_alternative(make_html_body(data), subtype='html')
    message.add_attachment(
        excel_bytes,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=excel_filename,
    )
    return message.as_bytes()
