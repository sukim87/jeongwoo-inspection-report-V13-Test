from datetime import date
from io import BytesIO
from pathlib import Path
import base64
import hashlib
import textwrap

import streamlit as st
from PIL import Image, ImageOps

from modules.application_reader import read_application
from modules.report_text import make_draft
from modules.excel_writer import build_report
from modules.email_helper import build_eml, make_html_body, make_subject

VERSION = "V13"
BASE = Path(__file__).resolve().parent

st.set_page_config(
    page_title="Jeongwoo AI 품질검사 결과서",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

INSPECTION_CODES = [
    ("서류검토", "DR"), ("열처리", "HT"), ("재료검사", "MI"),
    ("조립검사", "AS"), ("마킹검사", "MK"), ("압력시험", "PR"),
    ("취부검사", "FU"), ("성능시험", "PF"), ("치수검사", "DI"),
    ("비파괴검사", "NDE"), ("외관검사", "VI"), ("도장/도금", "PT"),
    ("가공검사", "MA"), ("포장/출하", "PK"), ("성형검사", "FM"),
    ("용접검사", "WD"),
]

INSPECTOR_NAMES = [
    "김상민", "김선웅", "김희성", "박재암", "송민준", "안병영",
    "안예지", "유병조", "이승현", "이진원", "이진표",
]


def html(content):
    st.markdown(textwrap.dedent(content).strip(), unsafe_allow_html=True)


def image_data_uri(path):
    path = Path(path)
    if not path.exists():
        return ""
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return "data:image/jpeg;base64," + encoded


hero_uri = image_data_uri(BASE / "assets" / "company_view_day.jpg")

html(f"""
<style>
:root{{--jw-red:#c8102e;--jw-text:#17233b}}
html,body,[class*="css"],[data-testid="stAppViewContainer"],[data-testid="stSidebar"]{{font-family:"Malgun Gothic","맑은 고딕",Arial,sans-serif!important;color:var(--jw-text)}}
[data-testid="stAppViewContainer"]{{background:#fff}}
[data-testid="stHeader"]{{background:rgba(255,255,255,.93)}}
.block-container{{max-width:1280px;padding-top:1.7rem!important;padding-bottom:4rem!important}}
section[data-testid="stSidebar"]{{background:linear-gradient(180deg,#f5f7fa 0%,#eef2f7 100%);border-right:1px solid #dbe1e9}}
section[data-testid="stSidebar"] .block-container{{padding:1.15rem .85rem!important}}
[data-testid="stSidebar"] img{{background:#fff;border-radius:12px;padding:8px;border:1px solid #e2e6ec}}
.side-card{{background:#fff;border:1px solid #dce2e9;border-top:4px solid var(--jw-red);border-radius:10px;padding:15px 14px;line-height:1.65;font-size:.86rem;box-shadow:0 4px 14px rgba(28,42,67,.05)}}
.side-footer{{margin-top:30px;padding:17px 6px 4px;border-top:1px solid #cfd6df;color:#687386;font-size:.72rem;line-height:1.6}}
.version-badge{{display:inline-block;margin:8px 0 10px;padding:5px 12px;border-radius:18px;background:#667085;color:#fff;font-weight:700}}
.hero{{position:relative;overflow:hidden;min-height:280px;border-radius:15px;background:linear-gradient(90deg,rgba(5,16,33,.96) 0%,rgba(5,16,33,.82) 48%,rgba(5,16,33,.12) 100%),url("{hero_uri}") center 58%/cover no-repeat;display:flex;align-items:center;padding:35px 38px;box-shadow:0 10px 28px rgba(20,35,60,.16)}}
.hero-copy{{max-width:720px;color:#fff}}
.hero-title{{font-size:3rem;line-height:1.23;font-weight:900;letter-spacing:-.055em;margin:0 0 18px;text-shadow:0 2px 8px rgba(0,0,0,.28)}}
.hero-title .accent{{color:#ff334f}}
.hero-sub{{font-size:1.02rem;line-height:1.75;font-weight:600;text-shadow:0 1px 5px rgba(0,0,0,.35)}}
.section-title{{display:flex;align-items:center;gap:10px;margin:18px 0 0;padding:13px 17px;background:linear-gradient(180deg,#f3f1e9,#e4e0d2);border:1px solid #d3cdbc;border-radius:11px 11px 0 0;color:#13213c;font-size:1.12rem;font-weight:900}}
[data-testid="stVerticalBlockBorderWrapper"]{{border-color:#e0e5ec!important;border-radius:0 0 12px 12px!important;box-shadow:0 4px 15px rgba(20,35,60,.035)}}
[data-baseweb="input"]>div,[data-baseweb="textarea"]>div,[data-baseweb="select"]>div{{background:#f7f9fc!important;border-color:#dde3ec!important;border-radius:8px!important}}
.stButton button,.stDownloadButton button{{border-radius:9px!important;font-weight:800!important;min-height:45px}}
.stButton button[kind="primary"],.stDownloadButton button[kind="primary"]{{background:linear-gradient(90deg,#c8102e,#ec153e)!important;border:0!important;box-shadow:0 6px 14px rgba(200,16,46,.18)}}
[data-testid="stFileUploaderDropzone"]{{background:#f8fafc!important;border:1.5px dashed #c9d2df!important;border-radius:10px!important}}
.quick-note{{font-size:.81rem;color:#667085;line-height:1.6;background:#f7f9fc;border:1px solid #e0e5ec;border-radius:8px;padding:10px 12px}}
.mailbox{{background:#fff;border:2px solid var(--jw-red);border-radius:12px;padding:16px;margin-top:18px}}
@media(max-width:900px){{.hero{{min-height:240px;padding:28px 24px;background-position:62% center}}.hero-title{{font-size:2.2rem}}.hero-sub{{font-size:.92rem}}}}
</style>
""")

with st.sidebar:
    logo_path = BASE / "assets" / "jeongwoo_logo.jpg"
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)

    html("""
    <div class="side-card">
      <b>☎ 시스템 문의 및 지원</b><br><br>
      <b>부서:</b> PS품질팀<br>
      <b>담당자:</b> 김선웅<br>
      <b>이메일:</b> <span style="color:#0067b8">sukim1@jimc.co.kr</span><br><br>
      <span style="color:#667085">프로그램 오류, 기능 개선,<br>검사양식 및 입력항목 변경 요청</span>
    </div>
    """)

    html(f"""
    <div class="side-footer">
      <b style="color:#17233b;font-size:.8rem">Jeongwoo AI Inspection Report</b><br>
      <span class="version-badge">Version {VERSION}</span><br>
      <a href="https://www.jimc.co.kr" target="_blank" style="color:#c8102e;font-weight:800;text-decoration:none">JIMC.CO.KR ↗</a><br><br>
      Copyright © JEONGWOO.<br>All Rights Reserved.
    </div>
    """)

html("""
<div class="hero">
  <div class="hero-copy">
    <div class="hero-title"><span class="accent">AI로 더 빠르게</span> 작성하고<br><span class="accent">공유하는</span> 품질검사 결과서</div>
    <div class="hero-sub">사진과 간략하게 검사내용을 입력하고 결과만 확인하세요.<br>입력 오타를 다듬어 검사 결과서를 작성하고 메일로 간편하게 공유할 수 있습니다.</div>
  </div>
</div>
""")

left_col, right_col = st.columns(2, gap="medium")
with left_col:
    with st.expander("👁 완성 결과서 예시 보기"):
        sample_path = BASE / "assets" / "sample_result.png"
        if sample_path.exists():
            st.image(str(sample_path), use_container_width=True)
with right_col:
    manual_candidates = [
        BASE / "assets" / "Jeongwoo_AI_품질검사_결과서_V12_사용매뉴얼.pdf",
        BASE / "assets" / "Jeongwoo_AI_품질검사_결과서_V11.1_사용매뉴얼.pdf",
    ]
    manual_path = next((p for p in manual_candidates if p.exists()), None)
    if manual_path:
        st.download_button(
            "📕 사용 매뉴얼 PDF 다운로드",
            data=manual_path.read_bytes(),
            file_name=manual_path.name,
            mime="application/pdf",
            use_container_width=True,
        )

html('<div class="section-title">1. 검사신청서 불러오기</div>')
with st.container(border=True):
    application_file = st.file_uploader(
        "검사신청서 PDF 또는 Excel 파일을 업로드하세요.",
        type=["pdf", "xlsx"],
        help="기존 또는 신규 검사신청서를 첨부하면 확실한 기본정보만 자동입력합니다.",
    )
    if application_file:
        file_signature = hashlib.sha256(application_file.getvalue()).hexdigest()
        if st.session_state.get("application_signature") != file_signature:
            try:
                application_info, application_preview = read_application(application_file)
                for key in ("customer", "part_name", "order_no", "po_step_no", "inspection_location"):
                    st.session_state[key] = application_info.get(key, "")
                st.session_state.application_signature = file_signature
                st.session_state.application_preview = application_preview
                st.session_state.application_info = application_info
            except Exception as error:
                st.error(f"신청서 분석 중 오류가 발생했습니다: {error}")

        application_info = st.session_state.get("application_info", {})
        review_fields = application_info.get("review_fields", [])
        field_names = {
            "customer": "고객사",
            "part_name": "부품명",
            "order_no": "수주번호",
            "po_step_no": "발주/공정번호",
            "inspection_location": "검사장소",
        }
        if review_fields:
            review_text = ", ".join(field_names.get(field, field) for field in review_fields)
            st.warning(f"확실한 정보만 자동입력했습니다. 직접 확인 필요: {review_text}")
        else:
            st.success("신청서 기본정보를 자동입력했습니다. 아래 내용을 최종 확인해 주세요.")

        preview = st.session_state.get("application_preview")
        if preview is not None:
            with st.expander("방향 보정된 신청서 미리보기"):
                st.image(preview, use_container_width=True)

html('<div class="section-title">2. 기본정보</div>')
with st.container(border=True):
    col1, col2 = st.columns(2)
    customer = col1.text_input("고객사", key="customer", placeholder="예: 한화오션")
    part_name = col2.text_input("부품명 *", key="part_name", placeholder="예: Cold Support")

    col1, col2 = st.columns(2)
    order_no = col1.text_input("수주번호 *", key="order_no", placeholder="예: H260926UD")
    po_step_no = col2.text_input("발주/공정번호 *", key="po_step_no", placeholder="예: PS2610022")

    col1, col2 = st.columns(2)
    inspection_location = col1.text_input(
        "검사장소 *",
        key="inspection_location",
        placeholder="신청서 업체/팀명이 있으면 자동입력",
    )
    inspector = col2.selectbox(
        "검사자 *",
        INSPECTOR_NAMES,
        index=INSPECTOR_NAMES.index("김선웅"),
        help="실제 검사를 수행한 검사원을 선택하세요.",
    )

    col1, col2 = st.columns(2)
    document_sequence = col1.number_input(
        "문서 순번",
        min_value=1,
        max_value=999,
        value=1,
        step=1,
        help="최초 001, 동일 건의 추가 검사·재검사는 002, 003 순으로 사용합니다.",
    )
    document_no = f"IRR-{order_no or 'ORDER'}-{po_step_no or 'PO-STEP'}-{int(document_sequence):03d}"
    col2.text_input("문서번호", value=document_no, disabled=True)

html('<div class="section-title">3. 검사종류 <span style="font-size:.8rem;font-weight:600;color:#667085">(복수 선택 가능)</span></div>')
with st.container(border=True):
    inspection_columns = st.columns(4)
    selected_inspections = []
    for index, (inspection_name, inspection_code) in enumerate(INSPECTION_CODES):
        label = f"{inspection_name} ({inspection_code})"
        if inspection_columns[index % 4].checkbox(label, key=f"inspection_{inspection_code}"):
            selected_inspections.append(label)

html('<div class="section-title">4. 검사 기준 및 검사내용</div>')
with st.container(border=True):
    selected_references = []
    reference_columns = st.columns(3)
    for column, reference_name in zip(reference_columns, ("도면", "구매사양서", "구매발주서")):
        if column.checkbox(reference_name, key=f"reference_{reference_name}"):
            selected_references.append(reference_name)

    manual_specification = st.text_input(
        "추가 검사기준 직접 입력 (선택)",
        placeholder="예: ASME B31.1 / DFT Min. 320 μm / 승인도면 Rev.2",
    )
    spec_parts = selected_references.copy()
    if manual_specification.strip():
        spec_parts.append(manual_specification.strip())
    specification = " / ".join(spec_parts)

    col1, col2 = st.columns(2)
    col1.text_area("검사기준 *", value=specification, height=95, disabled=True)
    inspection_content = col2.text_area(
        "검사내용 *",
        key="inspection_content",
        height=95,
        placeholder="현장에서 확인한 내용을 간단히 입력하세요.\n예: 치수 및 외관 스펙 만족",
    )
    judgment = st.selectbox("최종판정", ["합격", "불합격", "조건부 합격", "N/A"])

html('<div class="section-title">5. 검사사진 첨부</div>')
with st.container(border=True):
    photos = st.file_uploader(
        "검사사진을 업로드하세요. (최대 8장)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
    )
    rotations = []
    fit_modes = []
    if len(photos) > 8:
        st.warning("앞에서부터 8장까지만 자동 배치됩니다. 추가 사진은 을지 양식을 복사·편집해 주세요.")

    if photos:
        photo_columns = st.columns(2)
        for index, photo in enumerate(photos[:8]):
            with photo_columns[index % 2]:
                preview_image = ImageOps.exif_transpose(Image.open(BytesIO(photo.getvalue())))
                rotation_option = st.selectbox(
                    f"사진 {index + 1} 방향",
                    ["자동", "오른쪽 90°", "180°", "왼쪽 90°"],
                    key=f"photo_rotation_{index}",
                )
                rotation_degree = {
                    "자동": 0,
                    "오른쪽 90°": 90,
                    "180°": 180,
                    "왼쪽 90°": 270,
                }[rotation_option]
                rotations.append(rotation_degree)

                fit_mode = st.selectbox(
                    f"사진 {index + 1} 맞춤",
                    ["전체 사진 표시", "사진 칸 채우기"],
                    key=f"photo_fit_{index}",
                )
                fit_modes.append(fit_mode)

                if rotation_degree:
                    preview_image = preview_image.rotate(-rotation_degree, expand=True)
                st.image(preview_image, caption=f"사진 {index + 1}", use_container_width=True)

    html('<div class="quick-note">전체 사진 표시가 기본값이며 사진을 자르지 않습니다. 갑지에는 사진 1~2, 사진 3~8은 을지에 배치됩니다.</div>')

signature_source = "|".join(selected_inspections) + "|" + specification + "|" + inspection_content + "|" + judgment + "|" + str(len(photos))
draft_signature = hashlib.sha256(signature_source.encode("utf-8")).hexdigest()

if st.session_state.get("draft_signature") not in (None, draft_signature):
    st.session_state.draft_ready = False

if st.button("✨ 오타 교정 후 AI 초안 작성", type="primary", use_container_width=True):
    if not selected_inspections or not specification or not inspection_content.strip():
        st.error("검사종류, 검사기준, 검사내용을 먼저 입력해 주세요.")
    else:
        with st.spinner("AI가 검사내용을 정리하고 있습니다..."):
            ai_result, ai_action, ai_caption, corrected_spec, corrected_content = make_draft(
                selected_inspections,
                specification,
                inspection_content,
                judgment,
                len(photos),
            )
        st.session_state.ai_result = ai_result
        st.session_state.ai_action = ai_action
        st.session_state.ai_caption = ai_caption
        st.session_state.corrected_spec = corrected_spec
        st.session_state.corrected_content = corrected_content
        st.session_state.draft_ready = True
        st.session_state.draft_signature = draft_signature
        st.rerun()

if st.session_state.get("draft_ready"):
    html('<div class="section-title">6. AI 작성 결과 확인</div>')
    with st.container(border=True):
        col1, col2 = st.columns(2)
        col1.text_area("교정된 검사기준", value=st.session_state.get("corrected_spec", ""), height=95, disabled=True)
        col2.text_area("교정된 검사내용", value=st.session_state.get("corrected_content", ""), height=95, disabled=True)
        result_description = st.text_area("검사결과 문장", key="ai_result", height=110)
        required_action = st.text_area("조치요구", key="ai_action", height=90)
        photo_caption = st.text_area("사진 설명", key="ai_caption", height=80)

        missing_fields = [
            name
            for name, value in [
                ("부품명", part_name),
                ("수주번호", order_no),
                ("발주/공정번호", po_step_no),
                ("검사장소", inspection_location),
            ]
            if not str(value).strip()
        ]
        if missing_fields:
            st.warning("필수 입력: " + ", ".join(missing_fields))

        if st.button(
            "✅ 검사결과서 생성",
            type="primary",
            use_container_width=True,
            disabled=bool(missing_fields),
        ):
            report_data = {
                "customer": customer,
                "part_name": part_name,
                "order_no": order_no,
                "po_step_no": po_step_no,
                "inspection_location": inspection_location,
                "inspector": inspector,
                "date": date.today(),
                "doc": document_no,
                "types": " / ".join(selected_inspections),
                "spec": st.session_state.get("corrected_spec", specification),
                "inspection_content": st.session_state.get("corrected_content", inspection_content),
                "judgment": judgment,
                "result": result_description,
                "action": required_action,
                "caption": photo_caption,
            }

            st.session_state.mail_data = {
                "order_no": order_no,
                "po_step_no": po_step_no,
                "part_name": part_name,
                "inspection_location": inspection_location,
                "types": " / ".join(selected_inspections),
                "judgment": judgment,
                "inspection_content": st.session_state.get("corrected_content", inspection_content),
                "document_no": document_no,
                "inspector": inspector,
            }

            try:
                generated_excel, generated_filename = build_report(
                    report_data,
                    photos[:8],
                    rotations,
                    fit_modes,
                )
                st.session_state.out = generated_excel
                st.session_state.name = generated_filename
                st.success("검사결과서가 완성되었습니다.")
            except Exception as error:
                st.error(f"검사결과서 생성 중 오류: {error}")

if st.session_state.get("out"):
    st.download_button(
        "📥 검사결과서 다운로드",
        data=st.session_state.out,
        file_name=st.session_state.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    html("""
    <div class="mailbox">
      <b>✉️ 검사결과 메일 작성</b><br>
      <span style="color:#667085">수신자만 입력하면 제목, 본문, Excel 첨부가 포함된 Outlook 메일 파일을 만듭니다.</span>
    </div>
    """)

    to_value = st.text_input("받는 사람 이메일", placeholder="여러 명은 세미콜론(;)으로 구분")
    cc_value = st.text_input("참조 이메일 (선택)")

    mail_data = dict(st.session_state.get("mail_data", {}))
    mail_data.update({
        "order_no": order_no,
        "po_step_no": po_step_no,
        "part_name": part_name,
        "inspection_location": inspection_location,
        "types": " / ".join(selected_inspections),
        "judgment": judgment,
        "inspection_content": st.session_state.get("corrected_content", inspection_content),
        "document_no": document_no,
        "inspector": inspector,
    })
    st.session_state.mail_data = mail_data

    st.text_input("메일 제목", value=make_subject(mail_data), disabled=True)
    with st.expander("메일 본문 미리보기", expanded=True):
        mail_preview_html = textwrap.dedent(make_html_body(mail_data)).strip()
        st.markdown(mail_preview_html, unsafe_allow_html=True)

    eml_bytes = build_eml(
        mail_data,
        st.session_state.out,
        st.session_state.name,
        to_value,
        cc_value,
    )
    st.download_button(
        "✉️ 첨부 포함 Outlook 메일 열기 (.eml)",
        data=eml_bytes,
        file_name=f"{document_no}_메일.eml",
        mime="message/rfc822",
        type="primary",
        use_container_width=True,
    )
