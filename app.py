
import json
import re
import uuid
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st


# =========================
# 기본 설정
# =========================
st.set_page_config(
    page_title="BigHoneySangkibu",
    page_icon="🍯",
    layout="wide",
)

APP_TITLE = "🍯 BigHoneySangkibu"
APP_SUBTITLE = "수행평가 기반 생기부 작성 도우미"


# =========================
# 유틸 함수
# =========================
def init_state():
    if "settings" not in st.session_state:
        st.session_state.settings = {
            "school_year": "2026",
            "semester": "1학기",
            "school_level": "중학교",
            "grade": "1",
            "subject": "과학",
            "target_bytes_min": 700,
            "target_bytes_max": 800,
            "rules": {
                "noun_ending": True,
                "no_name": True,
                "no_student_start": True,
                "avoid_exaggeration": True,
                "evidence_based": True,
                "remove_redundancy": True,
            },
        }

    if "students" not in st.session_state:
        st.session_state.students = pd.DataFrame(
            columns=["student_id", "학년", "반", "번호", "성명"]
        )

    if "assessments" not in st.session_state:
        st.session_state.assessments = []

    if "items" not in st.session_state:
        st.session_state.items = []

    if "records" not in st.session_state:
        # key = f"{student_id}::{item_id}"
        # value = {"level": "", "comment": ""}
        st.session_state.records = {}

    if "results" not in st.session_state:
        # key = student_id
        # value = {"material": "", "generated": "", "edited": "", "bytes": 0, "created_at": ""}
        st.session_state.results = {}


def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def make_id(prefix):
    return f"{prefix}_{uuid.uuid4().hex[:8]}"


def byte_count(text):
    return len(str(text).encode("utf-8"))


def split_class_number(value):
    """
    나이스의 '반/번호' 값을 반, 번호로 분리한다.
    예: 1/12, 1-12, 1 12, 1반 12번 등의 단순한 형태를 처리한다.
    """
    text = clean_text(value)
    if not text:
        return "", ""

    nums = re.findall(r"\d+", text)

    if len(nums) >= 2:
        return str(int(nums[0])), str(int(nums[1]))

    # 숫자가 하나만 있는 경우: 원문을 번호로만 판단하기 어려우므로 그대로 번호에 넣음
    if len(nums) == 1:
        return "", str(int(nums[0]))

    return "", ""


def find_header_row(raw_df):
    """
    나이스 파일처럼 1행 또는 중간행에 헤더가 있을 수 있으므로
    '성명'과 '학년' 또는 '반/번호'가 있는 행을 헤더로 판단한다.
    """
    for i in range(min(len(raw_df), 30)):
        row_values = [clean_text(v) for v in raw_df.iloc[i].tolist()]

        has_name = any(v in ["성명", "이름", "학생명"] for v in row_values)
        has_grade = "학년" in row_values or "학년도" in row_values
        has_class_number = any("반/번호" in v for v in row_values) or (
            "반" in row_values and "번호" in row_values
        )

        if has_name and (has_grade or has_class_number):
            return i

    return None


def normalize_columns(df):
    new_cols = []
    for c in df.columns:
        c = clean_text(c)
        c = c.replace("\n", "").replace(" ", "")
        new_cols.append(c)
    df.columns = new_cols
    return df


def parse_neis_excel(uploaded_file):
    """
    나이스 세특 다운로드 파일에서 학생 명단 추출.
    기본 목표:
    - 학년
    - 반/번호 -> 반, 번호
    - 성명 -> 성명
    학생개인번호는 저장하지 않는다.
    """
    xls = pd.ExcelFile(uploaded_file)

    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        header_row = find_header_row(raw)

        if header_row is None:
            continue

        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
        df = normalize_columns(df)
        df = df.dropna(how="all").copy()

        # 후보 열 찾기
        grade_col = None
        name_col = None
        class_number_col = None
        class_col = None
        number_col = None
        school_year_col = None
        semester_col = None
        subject_col = None

        for col in df.columns:
            if col == "학년도":
                school_year_col = col
            elif col == "학기":
                semester_col = col
            elif col == "학년":
                grade_col = col
            elif col in ["성명", "이름", "학생명"]:
                name_col = col
            elif "반/번호" in col:
                class_number_col = col
            elif col == "반":
                class_col = col
            elif col == "번호":
                number_col = col
            elif col == "과목명":
                subject_col = col

        if name_col is None:
            continue

        rows = []
        for _, row in df.iterrows():
            name = clean_text(row.get(name_col, ""))
            if not name or name in ["성명", "이름", "학생명"]:
                continue

            grade = clean_text(row.get(grade_col, "")) if grade_col else ""
            class_no = ""
            number = ""

            if class_number_col:
                class_no, number = split_class_number(row.get(class_number_col, ""))
            else:
                class_no = clean_text(row.get(class_col, "")) if class_col else ""
                number = clean_text(row.get(number_col, "")) if number_col else ""

            class_no = re.sub(r"\D", "", class_no)
            number = re.sub(r"\D", "", number)

            if not grade:
                grade = st.session_state.settings.get("grade", "")

            rows.append(
                {
                    "student_id": make_id("stu"),
                    "학년": grade,
                    "반": class_no,
                    "번호": number,
                    "성명": name,
                }
            )

        students = pd.DataFrame(rows)

        # 기본 설정 자동 반영
        if len(df) > 0:
            if school_year_col and len(df[school_year_col].dropna()):
                first_year = clean_text(df[school_year_col].dropna().iloc[0])
                if first_year:
                    st.session_state.settings["school_year"] = first_year
            if semester_col and len(df[semester_col].dropna()):
                first_semester = clean_text(df[semester_col].dropna().iloc[0])
                if first_semester:
                    st.session_state.settings["semester"] = first_semester
            if subject_col and len(df[subject_col].dropna()):
                first_subject = clean_text(df[subject_col].dropna().iloc[0])
                if first_subject:
                    st.session_state.settings["subject"] = first_subject

        if not students.empty:
            students = students.sort_values(
                ["학년", "반", "번호"], key=lambda s: s.astype(str)
            ).reset_index(drop=True)
            return students, sheet_name

    return pd.DataFrame(columns=["student_id", "학년", "반", "번호", "성명"]), None


def project_to_json():
    data = {
        "settings": st.session_state.settings,
        "students": st.session_state.students.to_dict(orient="records"),
        "assessments": st.session_state.assessments,
        "items": st.session_state.items,
        "records": st.session_state.records,
        "results": st.session_state.results,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "BigHoneySangkibu",
    }
    return json.dumps(data, ensure_ascii=False, indent=2)


def load_project_json(uploaded_file):
    data = json.load(uploaded_file)
    st.session_state.settings = data.get("settings", st.session_state.settings)
    st.session_state.students = pd.DataFrame(data.get("students", []))
    st.session_state.assessments = data.get("assessments", [])
    st.session_state.items = data.get("items", [])
    st.session_state.records = data.get("records", {})
    st.session_state.results = data.get("results", {})


def get_items_for_assessment(assessment_id):
    return [it for it in st.session_state.items if it.get("assessment_id", "") == assessment_id]


def get_assessment_name(assessment_id):
    for a in st.session_state.assessments:
        if a["assessment_id"] == assessment_id:
            return a["name"]
    return ""


def record_key(student_id, item_id):
    return f"{student_id}::{item_id}"


def get_record(student_id, item_id):
    return st.session_state.records.get(
        record_key(student_id, item_id), {"level": "", "comment": ""}
    )


def set_record(student_id, item_id, level="", comment=""):
    st.session_state.records[record_key(student_id, item_id)] = {
        "level": clean_text(level),
        "comment": clean_text(comment),
    }


def parse_rubric_text(levels_text, rubrics_text):
    """
    수준 코드와 문구를 단순 파싱.
    levels_text: A,B,C,D,E
    rubrics_text:
    A=문구
    B=문구
    또는
    A: 문구
    B: 문구
    """
    levels = [x.strip() for x in re.split(r"[,/|]", levels_text) if x.strip()]
    rubrics = {level: "" for level in levels}

    for line in rubrics_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if "=" in line:
            k, v = line.split("=", 1)
        elif ":" in line:
            k, v = line.split(":", 1)
        else:
            continue
        k, v = k.strip(), v.strip()
        if k:
            rubrics[k] = v
            if k not in levels:
                levels.append(k)

    return levels, rubrics


def build_student_material(student):
    """
    한 학생의 API 입력 자료 생성.
    학생 이름은 화면 확인용 제목에는 나오지만, API 프롬프트 본문에는 이름 없이 작성할 수 있게 구성한다.
    """
    student_id = student["student_id"]
    lines = []

    lines.append("[학생 평가 자료]")
    lines.append(
        f"- 학년/반/번호: {student.get('학년','')}학년 {student.get('반','')}반 {student.get('번호','')}번"
    )
    lines.append("")

    grouped = {}
    for a in st.session_state.assessments:
        grouped[a["assessment_id"]] = []

    for item in st.session_state.items:
        rec = get_record(student_id, item["item_id"])
        level = rec.get("level", "")
        comment = rec.get("comment", "")
        teacher_comment = ""

        if item["type"] in ["rubric", "rubric_plus"]:
            teacher_comment = item.get("rubrics", {}).get(level, "")

        if item["type"] == "rubric":
            if level or teacher_comment:
                grouped[item.get("assessment_id", "")].append(
                    f"- {item['name']}: 성취수준 {level} / 교사의 평가: {teacher_comment}"
                )

        elif item["type"] == "comment":
            if comment:
                grouped[item.get("assessment_id", "")].append(f"- {item['name']}: {comment}")

        elif item["type"] == "rubric_plus":
            if level or teacher_comment or comment:
                text = f"- {item['name']}: 성취수준 {level} / 교사의 평가: {teacher_comment}"
                if comment:
                    text += f" / 추가 코멘트: {comment}"
                grouped[item.get("assessment_id", "")].append(text)

    count = 1
    for a in st.session_state.assessments:
        chunks = grouped.get(a["assessment_id"], [])
        if not chunks:
            continue
        lines.append(f"{count}. {a['name']}")
        if a.get("area"):
            lines.append(f"- 영역/단원: {a.get('area','')}")
        if a.get("description"):
            lines.append(f"- 활동/성취기준: {a.get('description','')}")
        lines.extend(chunks)
        lines.append("")
        count += 1

    return "\n".join(lines).strip()


def build_prompt(material):
    s = st.session_state.settings
    rule_text = []
    if s["rules"].get("noun_ending"):
        rule_text.append("- 문장은 생기부 문체에 맞게 명사형 종결을 사용한다. 예: 분석함, 정리함, 제시함, 탐색함.")
    if s["rules"].get("no_name"):
        rule_text.append("- 학생 이름을 쓰지 않는다.")
    if s["rules"].get("no_student_start"):
        rule_text.append("- 첫 문장을 '학생은'으로 시작하지 않는다.")
    if s["rules"].get("avoid_exaggeration"):
        rule_text.append("- '깊은 이해', '창의융합', '혁신적' 같은 과장된 표현을 피한다.")
    if s["rules"].get("evidence_based"):
        rule_text.append("- 제공된 평가 자료에 근거한 내용만 쓴다.")
    if s["rules"].get("remove_redundancy"):
        rule_text.append("- 중복되는 평가 문구는 자연스럽게 통합한다.")

    prompt = f"""
너는 중학교 {s.get('subject','과학')} 교과 세부능력 및 특기사항을 작성하는 교사 보조 도구이다.

아래 학생 평가 자료를 바탕으로 교과 세부능력 및 특기사항을 작성하라.

[작성 조건]
- 목표 분량: {s.get('target_bytes_min',700)}~{s.get('target_bytes_max',800)} byte
{chr(10).join(rule_text)}
- 한 문단으로 작성한다.
- 문장 사이 연결을 자연스럽게 다듬는다.
- 평가 자료에 없는 성격, 인성, 태도, 진로 내용은 임의로 추가하지 않는다.

[학생 평가 자료]
{material}

[최종 출력]
세부능력 및 특기사항 문장만 출력하라.
""".strip()
    return prompt


def fallback_generate(material):
    """
    API 키가 없을 때 테스트용으로 쓰는 간단한 문장 조합기.
    실제 생기부 품질은 API 생성이 훨씬 낫다.
    """
    sentences = []

    for line in material.splitlines():
        line = line.strip()
        if not line.startswith("- "):
            continue
        if "교사의 평가:" in line:
            part = line.split("교사의 평가:", 1)[1].strip()
            if " / 추가 코멘트:" in part:
                base, extra = part.split(" / 추가 코멘트:", 1)
                if base:
                    sentences.append(base.strip())
                if extra:
                    sentences.append(extra.strip())
            elif part:
                sentences.append(part.strip())
        elif "):" not in line and ":" in line:
            part = line.split(":", 1)[1].strip()
            if part and "학년/반/번호" not in line:
                sentences.append(part)

    # 중복 제거
    seen = set()
    unique = []
    for s in sentences:
        s = s.strip().rstrip(".")
        if s and s not in seen:
            unique.append(s)
            seen.add(s)

    if not unique:
        return "입력된 평가 자료가 부족하여 생기부 문장을 생성하기 어려움."

    text = ". ".join(unique).strip()
    if not text.endswith(("함", "임", "음", ".")):
        text += "함"
    return text


def generate_with_openai(prompt, api_key, model):
    """
    OpenAI API를 이용한 생성.
    - api_key가 없으면 fallback_generate를 사용한다.
    - Streamlit Cloud에서는 Settings > Secrets에 OPENAI_API_KEY를 넣는 방식도 가능하다.
    """
    if not api_key:
        return None

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        response = client.responses.create(
            model=model,
            input=prompt,
        )
        return response.output_text.strip()

    except Exception as e:
        st.error(f"API 생성 중 오류가 발생했습니다: {e}")
        return None


def export_excel():
    students = st.session_state.students.copy()

    result_rows = []
    for _, stu in students.iterrows():
        sid = stu["student_id"]
        r = st.session_state.results.get(sid, {})
        result_rows.append(
            {
                "학년": stu.get("학년", ""),
                "반": stu.get("반", ""),
                "번호": stu.get("번호", ""),
                "성명": stu.get("성명", ""),
                "API 입력자료": r.get("material", ""),
                "생성 문구": r.get("generated", ""),
                "교사 수정 문구": r.get("edited", r.get("generated", "")),
                "byte": byte_count(r.get("edited", r.get("generated", ""))),
                "생성일시": r.get("created_at", ""),
            }
        )

    result_df = pd.DataFrame(result_rows)

    assessment_df = pd.DataFrame(st.session_state.assessments)
    item_rows = []
    for item in st.session_state.items:
        row = item.copy()
        row["assessment_name"] = get_assessment_name(item.get("assessment_id", ""))
        row["levels"] = ", ".join(item.get("levels", []))
        row["rubrics"] = "\n".join(
            [f"{k}: {v}" for k, v in item.get("rubrics", {}).items()]
        )
        item_rows.append(row)
    item_df = pd.DataFrame(item_rows)

    record_rows = []
    for _, stu in students.iterrows():
        for item in st.session_state.items:
            rec = get_record(stu["student_id"], item["item_id"])
            record_rows.append(
                {
                    "학년": stu.get("학년", ""),
                    "반": stu.get("반", ""),
                    "번호": stu.get("번호", ""),
                    "성명": stu.get("성명", ""),
                    "수행평가": get_assessment_name(item.get("assessment_id", "")),
                    "기록항목": item["name"],
                    "기록방식": item["type"],
                    "성취수준": rec.get("level", ""),
                    "개별코멘트": rec.get("comment", ""),
                }
            )
    record_df = pd.DataFrame(record_rows)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        students.drop(columns=["student_id"], errors="ignore").to_excel(
            writer, sheet_name="학생명단", index=False
        )
        assessment_df.to_excel(writer, sheet_name="수행평가", index=False)
        item_df.to_excel(writer, sheet_name="기록항목_루브릭", index=False)
        record_df.to_excel(writer, sheet_name="학생별기록", index=False)
        result_df.to_excel(writer, sheet_name="최종생기부", index=False)

        # 보기 좋게 열 너비 조정
        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            for col in ws.columns:
                max_len = 8
                col_letter = col[0].column_letter
                for cell in col:
                    value = clean_text(cell.value)
                    max_len = max(max_len, min(len(value), 50))
                ws.column_dimensions[col_letter].width = max_len + 2

    output.seek(0)
    return output


def load_sample_data():
    st.session_state.students = pd.DataFrame(
        [
            {"student_id": make_id("stu"), "학년": "1", "반": "1", "번호": "1", "성명": "강나은"},
            {"student_id": make_id("stu"), "학년": "1", "반": "1", "번호": "2", "성명": "경송혜"},
            {"student_id": make_id("stu"), "학년": "1", "반": "1", "번호": "3", "성명": "김보배"},
        ]
    )

    a1 = make_id("assess")
    a2 = make_id("assess")
    st.session_state.assessments = [
        {
            "assessment_id": a1,
            "name": "생태지도 만들기",
            "area": "생물과 환경",
            "description": "학교 운동장에서 생물을 관찰하고 생태지도를 제작하는 활동",
            "order": 1,
            "use": True,
        },
        {
            "assessment_id": a2,
            "name": "비열 탐구",
            "area": "열과 우리 생활",
            "description": "서로 다른 물체의 비열 차이를 비교하고 일상생활 사례와 연결하는 활동",
            "order": 2,
            "use": True,
        },
    ]

    item1 = make_id("item")
    item2 = make_id("item")
    item3 = make_id("item")

    st.session_state.items = [
        {
            "item_id": item1,
            "assessment_id": a1,
            "name": "생태지도 결과물 평가",
            "type": "rubric",
            "levels": ["A", "B", "C", "D", "E"],
            "rubrics": {
                "A": "학교 운동장에서 생물과 환경 요소를 세밀하게 관찰하고 생태지도를 완성도 있게 구성함",
                "B": "학교 운동장에서 생물을 관찰하고 생태지도에 적절히 나타냄",
                "C": "생태지도 작성 활동에 참여하였으나 생물과 환경 요소의 관계 표현이 다소 부족함",
                "D": "생태지도 작성 활동에 참여하였으나 관찰 결과 정리가 부족함",
                "E": "생태지도 작성 활동에서 보완이 필요함",
            },
            "order": 1,
        },
        {
            "item_id": item2,
            "assessment_id": a1,
            "name": "생태지도 개별 관찰 내용",
            "type": "comment",
            "levels": [],
            "rubrics": {},
            "order": 2,
        },
        {
            "item_id": item3,
            "assessment_id": a2,
            "name": "비열 개념 적용",
            "type": "rubric_plus",
            "levels": ["A", "B", "C", "D", "E"],
            "rubrics": {
                "A": "비열 차이에 따른 온도 변화의 차이를 근거를 들어 설명함",
                "B": "비열 차이와 온도 변화의 관계를 대체로 설명함",
                "C": "비열의 의미를 일부 이해하였으나 온도 변화와의 관계 설명이 다소 부족함",
                "D": "비열과 온도 변화의 관계를 설명하는 데 보완이 필요함",
                "E": "비열 개념 적용 활동에서 지속적인 보완이 필요함",
            },
            "order": 1,
        },
    ]

    st.session_state.records = {}
    for _, stu in st.session_state.students.iterrows():
        set_record(stu["student_id"], item1, "A", "")
        set_record(stu["student_id"], item2, "", "운동장 가장자리의 식물과 곤충을 관찰하고 특징을 정리함")
        set_record(stu["student_id"], item3, "A", "나무와 금속의 온도 변화 차이를 사례와 연결함")

    st.session_state.results = {}


# =========================
# 앱 시작
# =========================
init_state()

st.title(APP_TITLE)
st.caption(APP_SUBTITLE)

with st.sidebar:
    st.header("작업 관리")

    uploaded_project = st.file_uploader(
        "프로젝트 JSON 불러오기",
        type=["json"],
        help="이전에 저장한 BigHoneySangkibu 프로젝트 파일을 다시 불러옵니다.",
    )
    if uploaded_project and st.button("프로젝트 불러오기"):
        load_project_json(uploaded_project)
        st.success("프로젝트를 불러왔습니다.")
        st.rerun()

    st.download_button(
        "현재 프로젝트 JSON 저장",
        data=project_to_json(),
        file_name="BigHoneySangkibu_project.json",
        mime="application/json",
    )

    st.divider()

    if st.button("샘플 데이터 불러오기"):
        load_sample_data()
        st.success("샘플 데이터를 불러왔습니다.")
        st.rerun()

    if st.button("전체 초기화", type="secondary"):
        for key in ["settings", "students", "assessments", "items", "records", "results"]:
            if key in st.session_state:
                del st.session_state[key]
        init_state()
        st.success("초기화했습니다.")
        st.rerun()


tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "① 기본 설정",
        "② 학생 명단 업로드",
        "③ 수행평가 설계",
        "④ 학생별 기록 입력",
        "⑤ API 자료 확인",
        "⑥ 생기부 생성/다운로드",
    ]
)


# =========================
# ① 기본 설정
# =========================
with tab1:
    st.subheader("① 기본 설정")

    s = st.session_state.settings

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        s["school_year"] = st.text_input("학년도", value=s.get("school_year", "2026"))
    with c2:
        s["semester"] = st.selectbox(
            "학기",
            ["1학기", "2학기"],
            index=0 if s.get("semester") != "2학기" else 1,
        )
    with c3:
        s["school_level"] = st.selectbox(
            "학교급",
            ["중학교", "고등학교", "초등학교"],
            index=["중학교", "고등학교", "초등학교"].index(s.get("school_level", "중학교")),
        )
    with c4:
        s["grade"] = st.text_input("학년", value=s.get("grade", "1"))

    c5, c6, c7 = st.columns(3)
    with c5:
        s["subject"] = st.text_input("과목명", value=s.get("subject", "과학"))
    with c6:
        s["target_bytes_min"] = st.number_input(
            "목표 최소 byte", min_value=100, max_value=2000, value=int(s.get("target_bytes_min", 700)), step=50
        )
    with c7:
        s["target_bytes_max"] = st.number_input(
            "목표 최대 byte", min_value=100, max_value=2000, value=int(s.get("target_bytes_max", 800)), step=50
        )

    st.markdown("#### 생기부 작성 규칙")
    r = s["rules"]
    c1, c2, c3 = st.columns(3)
    with c1:
        r["noun_ending"] = st.checkbox("명사형 종결 사용", value=r.get("noun_ending", True))
        r["no_name"] = st.checkbox("학생 이름 제외", value=r.get("no_name", True))
    with c2:
        r["no_student_start"] = st.checkbox("'학생은'으로 시작하지 않기", value=r.get("no_student_start", True))
        r["avoid_exaggeration"] = st.checkbox("과장 표현 줄이기", value=r.get("avoid_exaggeration", True))
    with c3:
        r["evidence_based"] = st.checkbox("활동 근거 중심", value=r.get("evidence_based", True))
        r["remove_redundancy"] = st.checkbox("중복 문장 정리", value=r.get("remove_redundancy", True))

    st.info("1차 버전은 DB 없이 작동합니다. 작업을 이어서 하려면 왼쪽의 '현재 프로젝트 JSON 저장'을 눌러 파일로 저장하세요.")


# =========================
# ② 학생 명단 업로드
# =========================
with tab2:
    st.subheader("② 학생 명단 업로드")

    st.markdown(
        """
        나이스에서 내려받은 세특 시트의 다음 열을 자동으로 찾습니다.

        `학년도 | 학기 | 학년 | 반/번호 | 학생개인번호 | 성명 | 과목명 | 세부능력 및 특기사항`

        실제 명단에는 `학년`, `반/번호`, `성명`만 사용합니다. `학생개인번호`는 저장하지 않습니다.
        """
    )

    uploaded_excel = st.file_uploader("나이스 엑셀 파일 업로드", type=["xlsx", "xls"])

    if uploaded_excel:
        try:
            students, sheet_name = parse_neis_excel(uploaded_excel)
            if students.empty:
                st.error("학생 명단을 찾지 못했습니다. 헤더에 '학년', '반/번호', '성명'이 있는지 확인하세요.")
            else:
                st.success(f"'{sheet_name}' 시트에서 {len(students)}명의 학생을 찾았습니다.")
                st.dataframe(students.drop(columns=["student_id"]), use_container_width=True)

                if st.button("이 명단 사용하기"):
                    st.session_state.students = students
                    st.success("학생 명단을 적용했습니다.")
                    st.rerun()

        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")

    st.divider()
    st.markdown("#### 직접 입력 / 수정")
    st.caption("업로드 없이 테스트하거나, 업로드한 명단을 수정할 때 사용합니다.")

    editable_students = st.session_state.students.copy()
    if editable_students.empty:
        editable_students = pd.DataFrame(
            [{"student_id": make_id("stu"), "학년": st.session_state.settings["grade"], "반": "1", "번호": "1", "성명": ""}]
        )

    edited_students = st.data_editor(
        editable_students.drop(columns=["student_id"], errors="ignore"),
        num_rows="dynamic",
        use_container_width=True,
        key="students_editor",
    )

    if st.button("직접 입력 명단 저장"):
        new_df = edited_students.copy()
        if "student_id" not in new_df.columns:
            new_df.insert(0, "student_id", [make_id("stu") for _ in range(len(new_df))])
        new_df = new_df[new_df["성명"].astype(str).str.strip() != ""].reset_index(drop=True)
        st.session_state.students = new_df
        st.success("학생 명단을 저장했습니다.")
        st.rerun()


# =========================
# ③ 수행평가 설계
# =========================
with tab3:
    st.subheader("③ 수행평가 / 관찰 항목 설계")

    with st.expander("➕ 수행평가 추가", expanded=True):
        with st.form("add_assessment_form"):
            c1, c2 = st.columns(2)
            with c1:
                new_assessment_name = st.text_input("수행평가명", placeholder="예: 생태지도 만들기")
                new_area = st.text_input("영역/단원", placeholder="예: 생물과 환경")
            with c2:
                new_order = st.number_input("표시 순서", min_value=1, value=len(st.session_state.assessments) + 1, step=1)
                new_use = st.checkbox("사용", value=True)

            new_desc = st.text_area(
                "성취기준 / 활동 설명",
                placeholder="예: 학교 운동장에서 생물을 관찰하고 생태지도를 제작하는 활동",
                height=80,
            )

            submitted = st.form_submit_button("수행평가 추가")
            if submitted:
                if not new_assessment_name.strip():
                    st.warning("수행평가명을 입력하세요.")
                else:
                    st.session_state.assessments.append(
                        {
                            "assessment_id": make_id("assess"),
                            "name": new_assessment_name.strip(),
                            "area": new_area.strip(),
                            "description": new_desc.strip(),
                            "order": int(new_order),
                            "use": new_use,
                        }
                    )
                    st.success("수행평가를 추가했습니다.")
                    st.rerun()

    st.markdown("#### 등록된 수행평가와 기록 항목")

    if not st.session_state.assessments:
        st.info("아직 등록된 수행평가가 없습니다. 먼저 수행평가를 추가하세요.")

    for a in sorted(st.session_state.assessments, key=lambda x: x.get("order", 999)):
        with st.expander(f"📌 {a['name']}  /  {a.get('area','')}", expanded=False):
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                a["name"] = st.text_input("수행평가명 수정", value=a["name"], key=f"assess_name_{a['assessment_id']}")
                a["area"] = st.text_input("영역/단원 수정", value=a.get("area", ""), key=f"assess_area_{a['assessment_id']}")
            with c2:
                a["description"] = st.text_area("활동 설명 수정", value=a.get("description", ""), key=f"assess_desc_{a['assessment_id']}")
            with c3:
                a["order"] = st.number_input("순서", min_value=1, value=int(a.get("order", 1)), key=f"assess_order_{a['assessment_id']}")
                a["use"] = st.checkbox("사용", value=a.get("use", True), key=f"assess_use_{a['assessment_id']}")

            st.markdown("##### 기록 항목 추가")
            with st.form(f"add_item_form_{a['assessment_id']}"):
                item_name = st.text_input("기록 항목명", placeholder="예: 생태지도 결과물 평가")
                item_type_label = st.selectbox(
                    "기록 방식",
                    ["성취도 선택형", "개별 코멘트형", "성취도 + 추가 코멘트형"],
                )
                item_type_map = {
                    "성취도 선택형": "rubric",
                    "개별 코멘트형": "comment",
                    "성취도 + 추가 코멘트형": "rubric_plus",
                }

                levels_text = ""
                rubrics_text = ""

                if item_type_label != "개별 코멘트형":
                    levels_text = st.text_input("성취수준 코드", value="A,B,C,D,E")
                    rubrics_text = st.text_area(
                        "성취수준별 교사의 평가 문구",
                        value="A=우수한 수준으로 수행함\nB=대체로 적절하게 수행함\nC=일부 보완이 필요함\nD=기본적인 참여가 이루어짐\nE=지속적인 보완이 필요함",
                        height=140,
                        help="A=문구 또는 A:문구 형식으로 입력합니다.",
                    )

                item_order = st.number_input(
                    "항목 순서",
                    min_value=1,
                    value=len(get_items_for_assessment(a["assessment_id"])) + 1,
                    step=1,
                )

                add_item_submitted = st.form_submit_button("기록 항목 추가")

                if add_item_submitted:
                    if not item_name.strip():
                        st.warning("기록 항목명을 입력하세요.")
                    else:
                        if item_type_label == "개별 코멘트형":
                            levels, rubrics = [], {}
                        else:
                            levels, rubrics = parse_rubric_text(levels_text, rubrics_text)

                        st.session_state.items.append(
                            {
                                "item_id": make_id("item"),
                                "assessment_id": a["assessment_id"],
                                "name": item_name.strip(),
                                "type": item_type_map[item_type_label],
                                "levels": levels,
                                "rubrics": rubrics,
                                "order": int(item_order),
                            }
                        )
                        st.success("기록 항목을 추가했습니다.")
                        st.rerun()

            existing_items = sorted(get_items_for_assessment(a["assessment_id"]), key=lambda x: x.get("order", 999))
            if existing_items:
                st.markdown("##### 등록된 기록 항목")
                for item in existing_items:
                    type_kor = {
                        "rubric": "성취도 선택형",
                        "comment": "개별 코멘트형",
                        "rubric_plus": "성취도 + 추가 코멘트형",
                    }.get(item["type"], item["type"])

                    with st.container(border=True):
                        st.markdown(f"**{item['name']}** · {type_kor}")

                        c1, c2, c3 = st.columns([2, 1, 1])
                        with c1:
                            item["name"] = st.text_input(
                                "항목명",
                                value=item["name"],
                                key=f"item_name_{item['item_id']}",
                            )
                        with c2:
                            item["order"] = st.number_input(
                                "순서",
                                min_value=1,
                                value=int(item.get("order", 1)),
                                key=f"item_order_{item['item_id']}",
                            )
                        with c3:
                            if st.button("삭제", key=f"delete_item_{item['item_id']}"):
                                st.session_state.items = [
                                    x for x in st.session_state.items if x["item_id"] != item["item_id"]
                                ]
                                # 관련 학생 기록도 삭제
                                st.session_state.records = {
                                    k: v for k, v in st.session_state.records.items()
                                    if not k.endswith(f"::{item['item_id']}")
                                }
                                st.success("삭제했습니다.")
                                st.rerun()

                        if item["type"] in ["rubric", "rubric_plus"]:
                            levels_text2 = st.text_input(
                                "성취수준 코드",
                                value=", ".join(item.get("levels", [])),
                                key=f"item_levels_{item['item_id']}",
                            )
                            rubrics_text2 = st.text_area(
                                "루브릭 문구",
                                value="\n".join([f"{k}={v}" for k, v in item.get("rubrics", {}).items()]),
                                key=f"item_rubrics_{item['item_id']}",
                                height=120,
                            )
                            if st.button("루브릭 수정 저장", key=f"save_rubric_{item['item_id']}"):
                                levels, rubrics = parse_rubric_text(levels_text2, rubrics_text2)
                                item["levels"] = levels
                                item["rubrics"] = rubrics
                                st.success("루브릭을 저장했습니다.")
                                st.rerun()


# =========================
# ④ 학생별 기록 입력
# =========================
with tab4:
    st.subheader("④ 학생별 기록 입력")

    if st.session_state.students.empty:
        st.warning("먼저 학생 명단을 업로드하거나 입력하세요.")
    elif not st.session_state.items:
        st.warning("먼저 수행평가와 기록 항목을 추가하세요.")
    else:
        students = st.session_state.students.copy()
        assessments = [a for a in st.session_state.assessments if a.get("use", True)]
        assess_options = ["전체"] + [a["name"] for a in assessments]
        selected_assess_name = st.selectbox("수행평가 필터", assess_options)

        if selected_assess_name == "전체":
            selected_assess_ids = [a["assessment_id"] for a in assessments]
        else:
            selected_assess_ids = [a["assessment_id"] for a in assessments if a["name"] == selected_assess_name]

        selected_items = [
            it for it in st.session_state.items if it.get("assessment_id", "") in selected_assess_ids
        ]
        selected_items = sorted(selected_items, key=lambda x: (get_assessment_name(x["assessment_id"]), x.get("order", 999)))

        data_rows = []
        for _, stu in students.iterrows():
            row = {
                "student_id": stu["student_id"],
                "학년": stu.get("학년", ""),
                "반": stu.get("반", ""),
                "번호": stu.get("번호", ""),
                "성명": stu.get("성명", ""),
            }
            for item in selected_items:
                rec = get_record(stu["student_id"], item["item_id"])
                base_col = f"{get_assessment_name(item['assessment_id'])} - {item['name']}"

                if item["type"] == "rubric":
                    row[base_col] = rec.get("level", "")
                elif item["type"] == "comment":
                    row[base_col] = rec.get("comment", "")
                elif item["type"] == "rubric_plus":
                    row[f"{base_col} 성취수준"] = rec.get("level", "")
                    row[f"{base_col} 추가코멘트"] = rec.get("comment", "")

            data_rows.append(row)

        input_df = pd.DataFrame(data_rows)

        column_config = {}
        for item in selected_items:
            base_col = f"{get_assessment_name(item['assessment_id'])} - {item['name']}"
            if item["type"] == "rubric":
                column_config[base_col] = st.column_config.SelectboxColumn(
                    base_col,
                    options=[""] + item.get("levels", []),
                    required=False,
                )
            elif item["type"] == "rubric_plus":
                column_config[f"{base_col} 성취수준"] = st.column_config.SelectboxColumn(
                    f"{base_col} 성취수준",
                    options=[""] + item.get("levels", []),
                    required=False,
                )
                column_config[f"{base_col} 추가코멘트"] = st.column_config.TextColumn(
                    f"{base_col} 추가코멘트",
                    width="large",
                )
            else:
                column_config[base_col] = st.column_config.TextColumn(
                    base_col,
                    width="large",
                )

        edited_df = st.data_editor(
            input_df.drop(columns=["student_id"]),
            use_container_width=True,
            num_rows="fixed",
            disabled=["학년", "반", "번호", "성명"],
            column_config=column_config,
            key=f"record_editor_{selected_assess_name}",
            height=520,
        )

        if st.button("학생별 기록 저장"):
            edited_df = edited_df.copy()
            edited_df.insert(0, "student_id", input_df["student_id"].values)

            for _, row in edited_df.iterrows():
                sid = row["student_id"]
                for item in selected_items:
                    base_col = f"{get_assessment_name(item['assessment_id'])} - {item['name']}"

                    if item["type"] == "rubric":
                        set_record(sid, item["item_id"], level=row.get(base_col, ""), comment="")
                    elif item["type"] == "comment":
                        set_record(sid, item["item_id"], level="", comment=row.get(base_col, ""))
                    elif item["type"] == "rubric_plus":
                        set_record(
                            sid,
                            item["item_id"],
                            level=row.get(f"{base_col} 성취수준", ""),
                            comment=row.get(f"{base_col} 추가코멘트", ""),
                        )

            st.success("학생별 기록을 저장했습니다.")


# =========================
# ⑤ API 자료 확인
# =========================
with tab5:
    st.subheader("⑤ API 입력 자료 확인")

    if st.session_state.students.empty:
        st.warning("먼저 학생 명단을 입력하세요.")
    else:
        students = st.session_state.students.copy()
        label_map = {
            f"{row['반']}반 {row['번호']}번 {row['성명']}": row
            for _, row in students.iterrows()
        }

        selected_label = st.selectbox("학생 선택", list(label_map.keys()))
        selected_student = label_map[selected_label]
        material = build_student_material(selected_student)

        st.markdown("#### API 입력 자료 미리보기")
        st.text_area("이 내용이 API 프롬프트의 평가 자료로 들어갑니다.", value=material, height=420)

        with st.expander("실제 프롬프트 보기"):
            st.text_area("프롬프트", value=build_prompt(material), height=420)


# =========================
# ⑥ 생기부 생성/다운로드
# =========================
with tab6:
    st.subheader("⑥ 생기부 생성 / 수정 / 다운로드")

    if st.session_state.students.empty:
        st.warning("먼저 학생 명단을 입력하세요.")
    else:
        try:
            default_api_key = st.secrets.get("OPENAI_API_KEY", "")
        except Exception:
            default_api_key = ""

        c1, c2 = st.columns([2, 1])
        with c1:
            api_key = st.text_input(
                "OpenAI API Key",
                value=default_api_key,
                type="password",
                help="입력하지 않으면 API 없이 간단 조합 방식으로 생성됩니다. Streamlit Cloud에서는 Secrets에 OPENAI_API_KEY로 저장할 수 있습니다.",
            )
        with c2:
            model = st.text_input("모델명", value="gpt-5.5")

        students = st.session_state.students.copy()
        student_labels = {
            f"{row['반']}반 {row['번호']}번 {row['성명']}": row
            for _, row in students.iterrows()
        }
        selected_label = st.selectbox("생성할 학생", list(student_labels.keys()))
        selected_student = student_labels[selected_label]

        col_a, col_b = st.columns(2)

        with col_a:
            if st.button("선택 학생 생기부 생성", type="primary"):
                material = build_student_material(selected_student)
                prompt = build_prompt(material)

                generated = None
                if api_key:
                    with st.spinner("API로 생성 중..."):
                        generated = generate_with_openai(prompt, api_key, model)

                if not generated:
                    generated = fallback_generate(material)

                sid = selected_student["student_id"]
                st.session_state.results[sid] = {
                    "material": material,
                    "generated": generated,
                    "edited": generated,
                    "bytes": byte_count(generated),
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                st.success("생성했습니다.")
                st.rerun()

        with col_b:
            if st.button("전체 학생 생기부 생성"):
                progress = st.progress(0)
                total = len(students)

                for i, (_, stu) in enumerate(students.iterrows(), start=1):
                    material = build_student_material(stu)
                    prompt = build_prompt(material)

                    generated = None
                    if api_key:
                        generated = generate_with_openai(prompt, api_key, model)

                    if not generated:
                        generated = fallback_generate(material)

                    sid = stu["student_id"]
                    st.session_state.results[sid] = {
                        "material": material,
                        "generated": generated,
                        "edited": generated,
                        "bytes": byte_count(generated),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    progress.progress(i / total)

                st.success("전체 학생 생기부 생성을 완료했습니다.")
                st.rerun()

        st.divider()

        sid = selected_student["student_id"]
        result = st.session_state.results.get(sid)

        if result:
            st.markdown(f"#### {selected_label} 생성 결과")
            edited = st.text_area(
                "교사 수정 문구",
                value=result.get("edited", result.get("generated", "")),
                height=220,
            )
            b = byte_count(edited)
            st.caption(f"현재 byte: {b}")

            if b < st.session_state.settings["target_bytes_min"]:
                st.warning("목표 byte보다 짧습니다.")
            elif b > st.session_state.settings["target_bytes_max"]:
                st.warning("목표 byte보다 깁니다.")
            else:
                st.success("목표 byte 범위 안에 있습니다.")

            if st.button("수정 문구 저장"):
                result["edited"] = edited
                result["bytes"] = b
                st.session_state.results[sid] = result
                st.success("수정 문구를 저장했습니다.")
        else:
            st.info("아직 이 학생의 생기부 문구가 생성되지 않았습니다.")

        st.divider()
        st.markdown("#### 전체 결과표")

        result_rows = []
        for _, stu in students.iterrows():
            r = st.session_state.results.get(stu["student_id"], {})
            text = r.get("edited", r.get("generated", ""))
            result_rows.append(
                {
                    "학년": stu.get("학년", ""),
                    "반": stu.get("반", ""),
                    "번호": stu.get("번호", ""),
                    "성명": stu.get("성명", ""),
                    "생성/수정 문구": text,
                    "byte": byte_count(text),
                    "생성일시": r.get("created_at", ""),
                }
            )
        result_df = pd.DataFrame(result_rows)
        st.dataframe(result_df, use_container_width=True, height=320)

        excel_file = export_excel()
        st.download_button(
            "결과 엑셀 다운로드",
            data=excel_file,
            file_name=f"BigHoneySangkibu_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
