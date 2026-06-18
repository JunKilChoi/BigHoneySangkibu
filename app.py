import json
import re
import uuid
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

try:
    from streamlit_sortables import sort_items
except Exception:
    sort_items = None


# =========================
# 앱 기본 설정
# =========================
st.set_page_config(
    page_title="BigHoneySangkibu",
    page_icon="🍯",
    layout="wide",
)

APP_TITLE = "🍯 BigHoneySangkibu"
APP_SUBTITLE = "수행평가 기반 생기부 작성 도우미 · patched-20260618-v15-gray-drag-pretty-drag"


DEFAULT_RULES = """- 명사형 종결을 사용한다. 예: 분석함, 정리함, 제시함, 탐색함.
- 학생 이름을 쓰지 않는다.
- 첫 문장을 '학생은'으로 시작하지 않는다.
- '깊은 이해', '창의융합', '혁신적' 같은 과장 표현을 피한다.
- 제공된 평가 자료에 근거한 내용만 작성한다.
- 한 문단으로 작성하고, 중복되는 내용은 자연스럽게 통합한다.
- 평가 자료에 없는 인성, 진로, 성격, 태도 내용을 임의로 추가하지 않는다."""


# =========================
# 공통 유틸
# =========================
def make_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:8]}"


def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def byte_count(text: str) -> int:
    return len(clean_text(text).encode("utf-8"))


def to_int_or_big(value):
    text = re.sub(r"\D", "", clean_text(value))
    if text == "":
        return 999999
    return int(text)


def sort_students_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    result = df.copy()
    for col in ["student_id", "학년", "반", "번호", "성명"]:
        if col not in result.columns:
            result[col] = ""

    result["_학년정렬"] = result["학년"].map(to_int_or_big)
    result["_반정렬"] = result["반"].map(to_int_or_big)
    result["_번호정렬"] = result["번호"].map(to_int_or_big)

    result = result.sort_values(
        by=["_학년정렬", "_반정렬", "_번호정렬", "성명"],
        kind="stable",
    ).drop(columns=["_학년정렬", "_반정렬", "_번호정렬"])

    return result.reset_index(drop=True)



def needs_name_review(name) -> bool:
    """
    성명에 영어 알파벳이 포함된 경우에만 이름 확인 필요 학생으로 표시한다.
    블라인드 처리 후 생기는 * 문자는 확인 대상에 포함하지 않는다.
    """
    text = clean_text(name)
    if not text:
        return False

    return re.search(r"[A-Za-z]", text) is not None


def apply_name_review_edits(base_df: pd.DataFrame, review_df: pd.DataFrame) -> pd.DataFrame:
    """
    이름 확인 필요 학생만 따로 수정한 내용을 업로드 미리보기 명단에 반영한다.
    """
    if base_df.empty or review_df.empty or "student_id" not in base_df.columns or "student_id" not in review_df.columns:
        return base_df

    result = base_df.copy()
    review_map = review_df.set_index("student_id").to_dict(orient="index")

    for idx, row in result.iterrows():
        sid = row.get("student_id", "")
        if sid in review_map:
            for col in ["학년", "반", "번호", "성명"]:
                if col in review_map[sid]:
                    result.at[idx, col] = review_map[sid][col]

    return sort_students_df(result)



def mask_student_name(name, mask_char="*") -> str:
    """
    학생 이름을 개인정보 보호용으로 블라인드 처리한다.
    - 1글자: *
    - 2글자: 홍*
    - 3글자: 홍*동
    - 4글자 이상: 홍**동
    공백은 제거하지 않고 전체 문자열 기준으로 처리한다.
    """
    text = clean_text(name)
    if not text:
        return ""

    chars = list(text)
    length = len(chars)

    if length == 1:
        return mask_char
    if length == 2:
        return chars[0] + mask_char

    return chars[0] + (mask_char * (length - 2)) + chars[-1]


def mask_student_names_in_df(df: pd.DataFrame, mask_char="*") -> pd.DataFrame:
    result = df.copy()
    if "성명" not in result.columns:
        return result

    result["성명"] = result["성명"].apply(lambda x: mask_student_name(x, mask_char=mask_char))
    return result


def json_safe(obj):
    """
    프로젝트 JSON 저장 시 pandas/numpy 자료형 때문에 깨지지 않도록
    모든 값을 파이썬 기본 자료형 또는 문자열로 변환한다.
    """
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [json_safe(v) for v in obj]
    if isinstance(obj, tuple):
        return [json_safe(v) for v in obj]

    try:
        if pd.isna(obj):
            return ""
    except Exception:
        pass

    if hasattr(obj, "item"):
        try:
            return obj.item()
        except Exception:
            return str(obj)

    return obj


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
            "custom_rules": DEFAULT_RULES,
        }

    if "students" not in st.session_state:
        st.session_state.students = pd.DataFrame(
            columns=["student_id", "학년", "반", "번호", "성명"]
        )

    if "assessments" not in st.session_state:
        st.session_state.assessments = []

    # 주의: 세션 키 이름 items는 내부 메서드와 겹칠 수 있으므로 항상 대괄호 방식으로 접근한다.
    if "items" not in st.session_state:
        st.session_state["items"] = []

    if "records" not in st.session_state:
        st.session_state.records = {}

    if "results" not in st.session_state:
        st.session_state.results = {}

    if "generation_job" not in st.session_state:
        st.session_state.generation_job = {
            "active": False,
            "stop_requested": False,
            "student_ids": [],
            "index": 0,
            "log": [],
            "api_key": "",
            "model": "",
            "started_at": "",
            "finished_at": "",
        }


def sanitize_state():
    """
    이전 버전에서 만들어진 불완전한 세션 데이터가 남아도 앱이 죽지 않도록 보정한다.
    """
    settings = st.session_state.get("settings", {})
    if "custom_rules" not in settings:
        # 예전 체크박스 규칙을 사용하던 프로젝트도 새 방식으로 변환
        settings["custom_rules"] = DEFAULT_RULES
    if "target_bytes_min" not in settings:
        settings["target_bytes_min"] = 700
    if "target_bytes_max" not in settings:
        settings["target_bytes_max"] = 800
    st.session_state.settings = settings

    if not isinstance(st.session_state.get("students"), pd.DataFrame):
        st.session_state.students = pd.DataFrame(st.session_state.get("students", []))

    for col in ["student_id", "학년", "반", "번호", "성명"]:
        if col not in st.session_state.students.columns:
            st.session_state.students[col] = ""

    # 학생 ID가 비어 있으면 보정
    if not st.session_state.students.empty:
        st.session_state.students["student_id"] = st.session_state.students["student_id"].apply(
            lambda x: clean_text(x) if clean_text(x) else make_id("stu")
        )
        st.session_state.students = sort_students_df(st.session_state.students)

    # 수행평가 보정
    clean_assessments = []
    for idx, assessment in enumerate(st.session_state.get("assessments", []), start=1):
        if not isinstance(assessment, dict):
            continue
        if not assessment.get("assessment_id"):
            assessment["assessment_id"] = make_id("assess")
        assessment["name"] = assessment.get("name") or "이름 없는 수행평가"
        assessment["area"] = assessment.get("area", "")
        assessment["description"] = assessment.get("description", "")
        assessment["order"] = int(assessment.get("order", idx) or idx)
        assessment["use"] = bool(assessment.get("use", True))
        clean_assessments.append(assessment)

    st.session_state.assessments = clean_assessments
    valid_assessment_ids = {a["assessment_id"] for a in clean_assessments}

    # 평가 요소 보정
    clean_items = []
    for idx, item in enumerate(st.session_state.get("items", []), start=1):
        if not isinstance(item, dict):
            continue
        if item.get("assessment_id") not in valid_assessment_ids:
            continue
        if not item.get("item_id"):
            item["item_id"] = make_id("item")
        item["name"] = item.get("name") or "이름 없는 평가 요소"
        if item.get("type") not in ["rubric", "comment", "rubric_plus"]:
            item["type"] = "rubric"
        if not isinstance(item.get("levels", []), list):
            item["levels"] = []
        if not isinstance(item.get("rubrics", {}), dict):
            item["rubrics"] = {}
        if item["type"] == "comment":
            item["levels"] = []
            item["rubrics"] = {}
        item["order"] = int(item.get("order", idx) or idx)
        clean_items.append(item)

    st.session_state["items"] = clean_items
    valid_item_ids = {item["item_id"] for item in clean_items}

    # records 보정
    clean_records = {}
    for key, value in st.session_state.get("records", {}).items():
        key = str(key)
        if "::" not in key:
            continue
        item_id = key.split("::")[-1]
        if item_id not in valid_item_ids:
            continue
        if not isinstance(value, dict):
            continue
        clean_records[key] = {
            "level": clean_text(value.get("level", "")),
            "comment": clean_text(value.get("comment", "")),
        }
    st.session_state.records = clean_records


init_state()
sanitize_state()


# =========================
# 나이스 엑셀 파싱
# =========================
def normalize_col_name(col) -> str:
    return clean_text(col).replace(" ", "").replace("\n", "")


def split_class_number(value):
    text = clean_text(value)
    nums = re.findall(r"\d+", text)

    if len(nums) >= 2:
        return str(int(nums[0])), str(int(nums[1]))
    if len(nums) == 1:
        return "", str(int(nums[0]))
    return "", ""


def find_header_row(raw_df: pd.DataFrame):
    for i in range(min(len(raw_df), 30)):
        row_values = [normalize_col_name(v) for v in raw_df.iloc[i].tolist()]

        has_name = any(v in ["성명", "이름", "학생명"] for v in row_values)
        has_grade = any(v in ["학년", "학년도"] for v in row_values)
        has_class_number = any(v in ["반/번호", "반번호", "반"] for v in row_values)

        if has_name and (has_grade or has_class_number):
            return i

    return None


def parse_neis_excel(uploaded_file):
    """
    나이스 세특 다운로드 파일에서 학생 명단을 추출한다.
    여러 시트에 명단이 있으면 모두 합친다.
    """
    file_bytes = uploaded_file.getvalue()
    xls = pd.ExcelFile(BytesIO(file_bytes))

    all_students = []
    found_sheets = []

    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        header_row = find_header_row(raw)

        if header_row is None:
            continue

        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
        df.columns = [normalize_col_name(c) for c in df.columns]
        df = df.dropna(how="all").copy()

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
            elif col in ["반/번호", "반번호"]:
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
            if not grade:
                grade = st.session_state.settings.get("grade", "")

            if class_number_col:
                class_no, number = split_class_number(row.get(class_number_col, ""))
            else:
                class_no = clean_text(row.get(class_col, "")) if class_col else ""
                number = clean_text(row.get(number_col, "")) if number_col else ""

            class_no = re.sub(r"\D", "", clean_text(class_no))
            number = re.sub(r"\D", "", clean_text(number))

            rows.append(
                {
                    "student_id": make_id("stu"),
                    "학년": grade,
                    "반": class_no,
                    "번호": number,
                    "성명": name,
                }
            )

        if rows:
            found_sheets.append(sheet_name)
            all_students.extend(rows)

            # 기본 설정 자동 반영은 최초 발견값만 사용
            if school_year_col and not df[school_year_col].dropna().empty:
                first_year = clean_text(df[school_year_col].dropna().iloc[0])
                if first_year:
                    st.session_state.settings["school_year"] = first_year
            if semester_col and not df[semester_col].dropna().empty:
                first_semester = clean_text(df[semester_col].dropna().iloc[0])
                if first_semester:
                    st.session_state.settings["semester"] = first_semester
            if subject_col and not df[subject_col].dropna().empty:
                first_subject = clean_text(df[subject_col].dropna().iloc[0])
                if first_subject:
                    st.session_state.settings["subject"] = first_subject

    if not all_students:
        return pd.DataFrame(columns=["student_id", "학년", "반", "번호", "성명"]), []

    result = pd.DataFrame(all_students)
    result = sort_students_df(result)
    return result, found_sheets


def combine_students(existing_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    existing_df = existing_df.copy()
    new_df = new_df.copy()

    for df in [existing_df, new_df]:
        for col in ["student_id", "학년", "반", "번호", "성명"]:
            if col not in df.columns:
                df[col] = ""

    combined = pd.concat([existing_df, new_df], ignore_index=True)

    # 학년/반/번호/성명이 같은 학생은 중복 제거
    combined = combined.drop_duplicates(
        subset=["학년", "반", "번호", "성명"],
        keep="first",
    )

    return sort_students_df(combined)


# =========================
# 프로젝트 저장/불러오기
# =========================
def project_to_json() -> str:
    data = {
        "settings": st.session_state.settings,
        "students": st.session_state.students.to_dict(orient="records"),
        "assessments": st.session_state.assessments,
        "items": st.session_state["items"],
        "records": st.session_state.records,
        "results": st.session_state.results,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "BigHoneySangkibu",
        "version": "patched-20260618-v15-gray-drag-pretty-drag",
    }
    return json.dumps(json_safe(data), ensure_ascii=False, indent=2, default=str)


def load_project_json(uploaded_file):
    data = json.load(uploaded_file)
    st.session_state.settings = data.get("settings", st.session_state.settings)
    st.session_state.students = pd.DataFrame(data.get("students", []))
    st.session_state.assessments = data.get("assessments", [])
    st.session_state["items"] = data.get("items", [])
    st.session_state.records = data.get("records", {})
    st.session_state.results = data.get("results", {})
    sanitize_state()


# =========================
# 수행평가/기록항목 관련 함수
# =========================
def get_items_for_assessment(assessment_id):
    return [
        item
        for item in st.session_state["items"]
        if isinstance(item, dict) and item.get("assessment_id", "") == assessment_id
    ]



def normalize_item_orders(assessment_id):
    """
    한 수행평가 안의 평가 요소 순서를 1, 2, 3...처럼 중복 없이 다시 정리한다.
    이전 버전에서 같은 순서가 저장되어 있어도 자동으로 보정된다.
    """
    items = get_items_for_assessment(assessment_id)
    items = sorted(
        items,
        key=lambda x: (
            int(x.get("order", 999) or 999),
            clean_text(x.get("name", "")),
            clean_text(x.get("item_id", "")),
        ),
    )

    for idx, item in enumerate(items, start=1):
        item["order"] = idx



def normalize_assessment_orders():
    """
    수행평가 순서를 1, 2, 3...처럼 중복 없이 다시 정리한다.
    """
    assessments = sorted(
        st.session_state.assessments,
        key=lambda x: (
            int(x.get("order", 999) or 999),
            clean_text(x.get("name", "")),
            clean_text(x.get("assessment_id", "")),
        ),
    )

    for idx, assessment in enumerate(assessments, start=1):
        assessment["order"] = idx


def apply_assessment_drag_order(sorted_labels, label_to_assessment_id):
    """
    드래그 앤 드롭 결과에 맞춰 수행평가 order 값을 재배치한다.
    """
    id_to_assessment = {
        assessment.get("assessment_id", ""): assessment
        for assessment in st.session_state.assessments
    }

    for idx, label in enumerate(sorted_labels, start=1):
        assessment_id = label_to_assessment_id.get(label)
        if assessment_id in id_to_assessment:
            id_to_assessment[assessment_id]["order"] = idx


def apply_item_drag_order(assessment_id, sorted_labels, label_to_item_id):
    """
    드래그 앤 드롭 결과에 맞춰 한 수행평가 안의 평가 요소 order 값을 재배치한다.
    """
    id_to_item = {
        item.get("item_id", ""): item
        for item in get_items_for_assessment(assessment_id)
    }

    for idx, label in enumerate(sorted_labels, start=1):
        item_id = label_to_item_id.get(label)
        if item_id in id_to_item:
            id_to_item[item_id]["order"] = idx


def sortable_style():
    """
    streamlit-sortables 0.3.1 이상에서 사용하는 회색 계열 드래그 카드 디자인.
    글씨색, 배경색, 테두리색을 모두 명시해 기본 빨간색/흰 글씨 스타일이 끼어들지 않게 한다.
    """
    return """
    .sortable-component,
    .sortable-component.vertical {
        box-sizing: border-box !important;
        width: 100% !important;
        padding: 12px !important;
        border-radius: 14px !important;
        border: 1px solid #D1D5DB !important;
        background: #F9FAFB !important;
        box-shadow: 0 1px 2px rgba(17, 24, 39, 0.05) !important;
    }

    .sortable-container {
        box-sizing: border-box !important;
        border: 0 !important;
        border-radius: 12px !important;
        background: transparent !important;
        padding: 0 !important;
        margin: 0 !important;
    }

    .sortable-container-header {
        display: none !important;
        height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        border: 0 !important;
        background: transparent !important;
        color: #111827 !important;
    }

    .sortable-container-body,
    .sortable-container-boy {
        background: transparent !important;
        border: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
    }

    .sortable-item {
        position: relative !important;
        box-sizing: border-box !important;
        display: flex !important;
        align-items: center !important;
        min-height: 48px !important;
        margin: 8px 0 !important;
        padding: 12px 16px 12px 48px !important;
        border-radius: 12px !important;
        border: 1px solid #D1D5DB !important;
        background: #FFFFFF !important;
        color: #111827 !important;
        font-size: 15px !important;
        font-weight: 650 !important;
        line-height: 1.45 !important;
        letter-spacing: -0.01em !important;
        box-shadow: 0 1px 2px rgba(17, 24, 39, 0.05) !important;
        cursor: grab !important;
        user-select: none !important;
        transition: transform 0.12s ease, border-color 0.12s ease, box-shadow 0.12s ease, background 0.12s ease !important;
    }

    .sortable-item::before {
        content: "⋮⋮" !important;
        position: absolute !important;
        left: 17px !important;
        top: 50% !important;
        transform: translateY(-50%) !important;
        color: #9CA3AF !important;
        font-size: 18px !important;
        font-weight: 800 !important;
        letter-spacing: -4px !important;
    }

    .sortable-item:hover {
        transform: translateY(-1px) !important;
        border-color: #9CA3AF !important;
        background: #F3F4F6 !important;
        color: #111827 !important;
        box-shadow: 0 2px 8px rgba(17, 24, 39, 0.10) !important;
    }

    .sortable-item:active {
        cursor: grabbing !important;
        transform: scale(0.995) !important;
        border-color: #6B7280 !important;
        background: #E5E7EB !important;
        color: #111827 !important;
        box-shadow: 0 4px 12px rgba(17, 24, 39, 0.12) !important;
    }
    """


def sort_labels_with_pretty_box(labels, key, header="정렬"):
    """
    수행평가와 평가 요소에 공통으로 쓰는 드래그 정렬 함수.
    회색 카드 스타일은 streamlit-sortables==0.3.1 이상에서만 안정적으로 적용된다.
    """
    if sort_items is None:
        st.warning("드래그 정렬 기능을 사용하려면 requirements.txt에 streamlit-sortables==0.3.1을 추가해야 합니다.")
        return labels

    if len(labels) < 2:
        return labels

    try:
        return sort_items(
            labels,
            header=header,
            direction="vertical",
            custom_style=sortable_style(),
            key=key,
        )
    except TypeError:
        st.error(
            "현재 설치된 streamlit-sortables 버전이 custom_style을 지원하지 않습니다. "
            "requirements.txt의 streamlit-sortables를 streamlit-sortables==0.3.1로 바꾼 뒤, Streamlit Cloud에서 Reboot app을 눌러주세요."
        )
        return labels
    except Exception as e:
        st.error(f"드래그 정렬 컴포넌트 오류: {e}")
        return labels

def shift_item_orders_for_insert(assessment_id, insert_order):
    """
    새 평가 요소를 특정 순서에 끼워 넣을 때,
    기존 평가 요소들의 순서를 뒤로 밀어 중복을 막는다.
    """
    normalize_item_orders(assessment_id)

    for item in get_items_for_assessment(assessment_id):
        current_order = int(item.get("order", 999) or 999)
        if current_order >= int(insert_order):
            item["order"] = current_order + 1


def move_item_to_order(assessment_id, item_id, new_order):
    """
    기존 평가 요소의 순서를 바꿀 때,
    같은 수행평가 안의 다른 평가 요소 순서도 함께 재배치한다.
    """
    items = sorted(
        get_items_for_assessment(assessment_id),
        key=lambda x: int(x.get("order", 999) or 999),
    )

    moving_item = None
    remaining_items = []

    for item in items:
        if item.get("item_id", "") == item_id:
            moving_item = item
        else:
            remaining_items.append(item)

    if moving_item is None:
        return

    new_order = max(1, min(int(new_order), len(items)))
    remaining_items.insert(new_order - 1, moving_item)

    for idx, item in enumerate(remaining_items, start=1):
        item["order"] = idx


def get_assessment_name(assessment_id):
    for assessment in st.session_state.assessments:
        if assessment.get("assessment_id", "") == assessment_id:
            return assessment.get("name", "")
    return ""


def record_key(student_id, item_id):
    return f"{student_id}::{item_id}"


def get_record(student_id, item_id):
    return st.session_state.records.get(
        record_key(student_id, item_id),
        {"level": "", "comment": ""},
    )


def set_record(student_id, item_id, level="", comment=""):
    st.session_state.records[record_key(student_id, item_id)] = {
        "level": clean_text(level),
        "comment": clean_text(comment),
    }


def parse_rubric_text(levels_text, rubrics_text):
    levels = [x.strip() for x in re.split(r"[,/|]", levels_text) if x.strip()]
    rubrics = {level: "" for level in levels}

    for line in rubrics_text.splitlines():
        line = line.strip()
        if not line:
            continue

        if "=" in line:
            key, value = line.split("=", 1)
        elif ":" in line:
            key, value = line.split(":", 1)
        else:
            continue

        key = key.strip()
        value = value.strip()

        if key:
            rubrics[key] = value
            if key not in levels:
                levels.append(key)

    return levels, rubrics


def item_type_to_kor(item_type):
    return {
        "rubric": "성취도 선택형",
        "comment": "개별 코멘트형",
        "rubric_plus": "성취도 + 추가 코멘트형",
    }.get(item_type, "성취도 선택형")


def item_type_from_kor(label):
    return {
        "성취도 선택형": "rubric",
        "개별 코멘트형": "comment",
        "성취도 + 추가 코멘트형": "rubric_plus",
    }.get(label, "rubric")



def default_level_code(index):
    default_codes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    if 0 <= index < len(default_codes):
        return default_codes[index]
    return str(index + 1)


def render_rubric_input_block(prefix, current_levels=None, current_rubrics=None):
    """
    성취수준 개수를 1~10개 중 선택하고, 개수에 맞춰
    성취수준 코드와 교사의 평가 문구를 각각 입력받는다.
    """
    current_levels = current_levels if isinstance(current_levels, list) else []
    current_rubrics = current_rubrics if isinstance(current_rubrics, dict) else {}

    default_count = len(current_levels) if current_levels else 5
    default_count = max(1, min(10, int(default_count)))

    level_count = st.selectbox(
        "성취수준 개수",
        options=list(range(1, 11)),
        index=default_count - 1,
        key=f"{prefix}_level_count",
    )

    st.markdown("**성취수준 코드와 성취 수준별 교사의 평가 문구**")

    levels = []
    rubrics = {}

    for idx in range(level_count):
        existing_code = current_levels[idx] if idx < len(current_levels) else default_level_code(idx)
        existing_text = current_rubrics.get(existing_code, "")

        col_code, col_text = st.columns([1, 5])
        with col_code:
            code = st.text_input(
                f"{idx + 1}번 코드",
                value=existing_code,
                key=f"{prefix}_code_{idx}",
                label_visibility="collapsed",
                placeholder="A",
            )
        with col_text:
            comment = st.text_area(
                f"{idx + 1}번 평가 문구",
                value=existing_text,
                key=f"{prefix}_comment_{idx}",
                height=80,
                label_visibility="collapsed",
                placeholder="내용을 입력하세요. 교사가 평가한 내용을 입력해야 합니다.",
            )

        code = clean_text(code)
        if code:
            levels.append(code)
            rubrics[code] = clean_text(comment)

    return levels, rubrics


# =========================
# API 입력 자료 및 생성
# =========================
def build_student_material(student):
    student_id = student["student_id"]
    lines = []

    lines.append("[학생 평가 자료]")
    lines.append(
        f"- 학년/반/번호: {student.get('학년', '')}학년 "
        f"{student.get('반', '')}반 {student.get('번호', '')}번"
    )
    lines.append("")

    used_assessments = [a for a in st.session_state.assessments if a.get("use", True)]
    used_assessments = sorted(used_assessments, key=lambda x: int(x.get("order", 999) or 999))

    count = 1
    for assessment in used_assessments:
        assessment_id = assessment.get("assessment_id", "")
        chunks = []

        item_list = sorted(
            get_items_for_assessment(assessment_id),
            key=lambda x: int(x.get("order", 999) or 999),
        )

        for item in item_list:
            rec = get_record(student_id, item.get("item_id", ""))
            level = rec.get("level", "")
            comment = rec.get("comment", "")
            item_type = item.get("type", "rubric")
            teacher_comment = ""

            if item_type in ["rubric", "rubric_plus"]:
                teacher_comment = item.get("rubrics", {}).get(level, "")

            if item_type == "rubric":
                if level or teacher_comment:
                    chunks.append(
                        f"- {item.get('name', '')}: 성취수준 {level} / 교사의 평가: {teacher_comment}"
                    )

            elif item_type == "comment":
                if comment:
                    chunks.append(f"- {item.get('name', '')}: {comment}")

            elif item_type == "rubric_plus":
                if level or teacher_comment or comment:
                    text = f"- {item.get('name', '')}: 성취수준 {level} / 교사의 평가: {teacher_comment}"
                    if comment:
                        text += f" / 추가 코멘트: {comment}"
                    chunks.append(text)

        if chunks:
            lines.append(f"{count}. {assessment.get('name', '')}")
            if assessment.get("area"):
                lines.append(f"- 영역/단원: {assessment.get('area', '')}")
            if assessment.get("description"):
                lines.append(f"- 활동/성취기준: {assessment.get('description', '')}")
            lines.extend(chunks)
            lines.append("")
            count += 1

    return "\n".join(lines).strip()


def build_prompt(material):
    settings = st.session_state.settings
    custom_rules = clean_text(settings.get("custom_rules", DEFAULT_RULES))

    prompt = f"""
너는 {settings.get('school_level', '중학교')} {settings.get('subject', '과학')} 교과 세부능력 및 특기사항을 작성하는 교사 보조 도구이다.

아래 [선생님 작성 규칙]은 모든 학생의 생기부 작성에 최우선으로 적용한다.

[선생님 작성 규칙]
{custom_rules}

[작성 조건]
- 목표 분량: {settings.get('target_bytes_min', 700)}~{settings.get('target_bytes_max', 800)} byte
- 한 문단으로 작성한다.
- 문장 사이 연결을 자연스럽게 다듬는다.
- 평가 자료에 없는 내용을 임의로 추가하지 않는다.

[학생 평가 자료]
{material}

[최종 출력]
세부능력 및 특기사항 문장만 출력하라.
""".strip()
    return prompt


def fallback_generate(material):
    sentences = []

    for line in material.splitlines():
        line = line.strip()
        if not line.startswith("- "):
            continue

        if "교사의 평가:" in line:
            part = line.split("교사의 평가:", 1)[1].strip()

            if " / 추가 코멘트:" in part:
                base, extra = part.split(" / 추가 코멘트:", 1)
                if base.strip():
                    sentences.append(base.strip().rstrip("."))
                if extra.strip():
                    sentences.append(extra.strip().rstrip("."))
            elif part:
                sentences.append(part.strip().rstrip("."))

        elif (
            ":" in line
            and "학년/반/번호" not in line
            and "영역/단원" not in line
            and "활동/성취기준" not in line
        ):
            part = line.split(":", 1)[1].strip()
            if part:
                sentences.append(part.strip().rstrip("."))

    seen = set()
    unique = []
    for sentence in sentences:
        if sentence and sentence not in seen:
            unique.append(sentence)
            seen.add(sentence)

    if not unique:
        return "입력된 평가 자료가 부족하여 세부능력 및 특기사항 문장 생성이 어려움."

    text = ". ".join(unique).strip()
    if not text.endswith(("함", "임", "음", ".")):
        text += "함"
    return text


def generate_with_openai(prompt, api_key, model):
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
    for _, student in students.iterrows():
        sid = student["student_id"]
        result = st.session_state.results.get(sid, {})
        final_text = result.get("edited", result.get("generated", ""))

        result_rows.append(
            {
                "학년": student.get("학년", ""),
                "반": student.get("반", ""),
                "번호": student.get("번호", ""),
                "성명": student.get("성명", ""),
                "API 입력자료": result.get("material", ""),
                "생성 문구": result.get("generated", ""),
                "교사 수정 문구": final_text,
                "byte": byte_count(final_text),
                "생성일시": result.get("created_at", ""),
            }
        )

    result_df = pd.DataFrame(result_rows)

    item_rows = []
    for item in st.session_state["items"]:
        item_rows.append(
            {
                "수행평가": get_assessment_name(item.get("assessment_id", "")),
                "기록항목": item.get("name", ""),
                "기록방식": item_type_to_kor(item.get("type", "")),
                "성취수준": ", ".join(item.get("levels", [])),
                "루브릭": "\n".join([f"{k}: {v}" for k, v in item.get("rubrics", {}).items()]),
                "순서": item.get("order", ""),
            }
        )

    item_df = pd.DataFrame(item_rows)

    record_rows = []
    for _, student in students.iterrows():
        for item in st.session_state["items"]:
            rec = get_record(student["student_id"], item.get("item_id", ""))

            record_rows.append(
                {
                    "학년": student.get("학년", ""),
                    "반": student.get("반", ""),
                    "번호": student.get("번호", ""),
                    "성명": student.get("성명", ""),
                    "수행평가": get_assessment_name(item.get("assessment_id", "")),
                    "기록항목": item.get("name", ""),
                    "기록방식": item_type_to_kor(item.get("type", "")),
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
        pd.DataFrame(st.session_state.assessments).to_excel(
            writer, sheet_name="수행평가", index=False
        )
        item_df.to_excel(writer, sheet_name="평가요소_루브릭", index=False)
        record_df.to_excel(writer, sheet_name="학생별기록", index=False)
        result_df.to_excel(writer, sheet_name="최종생기부", index=False)

        for worksheet in writer.book.worksheets:
            for col in worksheet.columns:
                max_len = 8
                col_letter = col[0].column_letter
                for cell in col:
                    value = clean_text(cell.value)
                    max_len = max(max_len, min(len(value), 60))
                worksheet.column_dimensions[col_letter].width = max_len + 2

    output.seek(0)
    return output


def load_sample_data():
    st.session_state.students = pd.DataFrame(
        [
            {"student_id": make_id("stu"), "학년": "1", "반": "1", "번호": "1", "성명": "강나은"},
            {"student_id": make_id("stu"), "학년": "1", "반": "1", "번호": "2", "성명": "경송혜"},
            {"student_id": make_id("stu"), "학년": "1", "반": "2", "번호": "1", "성명": "김보배"},
        ]
    )
    st.session_state.students = sort_students_df(st.session_state.students)

    a1 = make_id("assess")
    a2 = make_id("assess")
    i1 = make_id("item")
    i2 = make_id("item")
    i3 = make_id("item")

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

    st.session_state["items"] = [
        {
            "item_id": i1,
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
            "item_id": i2,
            "assessment_id": a1,
            "name": "생태지도 개별 관찰 내용",
            "type": "comment",
            "levels": [],
            "rubrics": {},
            "order": 2,
        },
        {
            "item_id": i3,
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
    for _, student in st.session_state.students.iterrows():
        set_record(student["student_id"], i1, "A", "")
        set_record(student["student_id"], i2, "", "운동장 가장자리의 식물과 곤충을 관찰하고 특징을 정리함")
        set_record(student["student_id"], i3, "A", "나무와 금속의 온도 변화 차이를 사례와 연결함")

    st.session_state.results = {}
    sanitize_state()


# =========================
# 앱 UI
# =========================
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
        for key in ["settings", "students", "assessments", "items", "records", "results", "generation_job"]:
            if key in st.session_state:
                del st.session_state[key]
        init_state()
        sanitize_state()
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

    settings = st.session_state.settings

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        settings["school_year"] = st.text_input("학년도", value=clean_text(settings.get("school_year", "2026")))
    with col2:
        settings["semester"] = st.selectbox(
            "학기",
            ["1학기", "2학기"],
            index=1 if settings.get("semester") == "2학기" else 0,
        )
    with col3:
        school_options = ["중학교", "고등학교", "초등학교"]
        current_school = settings.get("school_level", "중학교")
        settings["school_level"] = st.selectbox(
            "학교급",
            school_options,
            index=school_options.index(current_school) if current_school in school_options else 0,
        )
    with col4:
        settings["grade"] = st.text_input("학년", value=clean_text(settings.get("grade", "1")))

    col5, col6, col7 = st.columns(3)
    with col5:
        settings["subject"] = st.text_input("과목명", value=clean_text(settings.get("subject", "과학")))
    with col6:
        settings["target_bytes_min"] = st.number_input(
            "목표 최소 byte",
            min_value=100,
            max_value=3000,
            value=int(settings.get("target_bytes_min", 700)),
            step=50,
        )
    with col7:
        settings["target_bytes_max"] = st.number_input(
            "목표 최대 byte",
            min_value=100,
            max_value=3000,
            value=int(settings.get("target_bytes_max", 800)),
            step=50,
        )

    st.markdown("#### 생기부 작성 규칙")
    st.caption("여기에 적은 규칙은 API 프롬프트의 맨 앞부분에 모든 학생 공통 규칙으로 들어갑니다.")

    settings["custom_rules"] = st.text_area(
        "공통 작성 규칙",
        value=settings.get("custom_rules", DEFAULT_RULES),
        height=220,
        help="예: 명사형 종결, 금지 표현, 학생 이름 제외, 분량, 문체, 과목 특성 등을 자유롭게 적으세요.",
    )

    with st.expander("API 프롬프트에 들어가는 규칙 예시 보기"):
        st.code(settings["custom_rules"], language="text")

    st.info("작업을 이어서 하려면 왼쪽의 '현재 프로젝트 JSON 저장'을 눌러 파일로 저장하세요.")


# =========================
# ② 학생 명단 업로드
# =========================
with tab2:
    st.subheader("② 학생 명단 업로드")

    st.markdown(
        """
        나이스 세특 파일에서 `학년`, `반/번호`, `성명`을 자동 추출합니다.  
        여러 반 파일을 한 번에 올리거나, 파일을 나중에 하나씩 추가해도 기존 명단에 누적할 수 있습니다.
        """
    )

    st.info("아래의 '업로드 파일 미리보기'는 아직 최종 명단이 아닙니다. 최종 명단은 하단의 '현재 학생 명단'이며, 그 표에서 직접 수정할 수 있습니다.")

    uploaded_excels = st.file_uploader(
        "나이스 엑셀 파일 업로드",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

    parsed_students = pd.DataFrame(columns=["student_id", "학년", "반", "번호", "성명"])
    parsed_info = []

    if uploaded_excels:
        parsed_frames = []

        for file in uploaded_excels:
            try:
                students_df, sheet_names = parse_neis_excel(file)
                if not students_df.empty:
                    students_df["원본파일"] = file.name
                    parsed_frames.append(students_df)
                    parsed_info.append(f"{file.name}: {len(students_df)}명 / 시트 {', '.join(sheet_names)}")
                else:
                    parsed_info.append(f"{file.name}: 학생 명단을 찾지 못함")
            except Exception as e:
                parsed_info.append(f"{file.name}: 오류 - {e}")

        if parsed_frames:
            parsed_students = pd.concat(parsed_frames, ignore_index=True)
            parsed_students = parsed_students.drop(columns=["원본파일"], errors="ignore")
            parsed_students = sort_students_df(parsed_students)

        for info in parsed_info:
            st.caption(info)

        if not parsed_students.empty:
            st.markdown("#### 업로드 파일 미리보기")
            st.caption("이 표는 업로드한 파일에서 읽어온 임시 명단입니다. 아래 버튼을 눌러야 현재 학생 명단에 반영됩니다.")

            st.dataframe(
                parsed_students.drop(columns=["student_id"], errors="ignore"),
                use_container_width=True,
                height=260,
            )

            review_students = parsed_students[parsed_students["성명"].map(needs_name_review)].copy()
            adjusted_parsed_students = parsed_students.copy()

            if not review_students.empty:
                st.markdown("#### 이름 확인 필요 학생")
                st.warning(
                    "성명에 한글 외 문자가 포함된 학생만 따로 모았습니다. "
                    "필요하면 여기서 이름을 수정한 뒤 아래의 추가/교체 버튼을 누르세요. "
                    "수정이 필요 없으면 그냥 넘어가도 됩니다."
                )

                edited_review = st.data_editor(
                    review_students,
                    use_container_width=True,
                    height=220,
                    num_rows="fixed",
                    disabled=["student_id"],
                    column_config={
                        "student_id": None,
                    },
                    key="name_review_editor",
                )

                adjusted_parsed_students = apply_name_review_edits(parsed_students, edited_review)
            else:
                st.success("한글 외 문자가 포함된 성명은 발견되지 않았습니다.")

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("업로드 명단을 현재 명단에 추가"):
                    st.session_state.students = combine_students(st.session_state.students, adjusted_parsed_students)
                    st.success("업로드 명단을 현재 학생 명단에 추가했습니다.")
                    st.rerun()

            with col_b:
                if st.button("현재 명단을 업로드 명단으로 교체"):
                    st.session_state.students = sort_students_df(adjusted_parsed_students)
                    st.success("현재 학생 명단을 업로드 명단으로 교체했습니다.")
                    st.rerun()

    st.divider()
    st.markdown("#### 현재 학생 명단")
    st.caption("이 표가 최종 명단입니다. 여기서 직접 수정한 뒤 '현재 학생 명단 저장'을 누르면 이후 모든 입력 화면에 반영됩니다. API 입력 전 개인정보 보호가 필요하면 저장 후 '이름 블라인드 처리'를 누르세요.")

    editable_students = st.session_state.students.copy()

    if editable_students.empty:
        editable_students = pd.DataFrame(
            [
                {
                    "student_id": make_id("stu"),
                    "학년": st.session_state.settings.get("grade", "1"),
                    "반": "1",
                    "번호": "1",
                    "성명": "",
                }
            ]
        )

    for col in ["student_id", "학년", "반", "번호", "성명"]:
        if col not in editable_students.columns:
            editable_students[col] = ""

    edited_students = st.data_editor(
        editable_students[["student_id", "학년", "반", "번호", "성명"]],
        num_rows="dynamic",
        use_container_width=True,
        height=360,
        key="final_students_editor",
        column_config={
            "student_id": None,
            "학년": st.column_config.TextColumn("학년", width="small"),
            "반": st.column_config.TextColumn("반", width="small"),
            "번호": st.column_config.TextColumn("번호", width="small"),
            "성명": st.column_config.TextColumn("성명", width="medium"),
        },
    )

    # 버튼은 왼쪽부터 '반/번호 순대로 다시 정렬 → 이름 블라인드 처리 → 저장' 순서로 붙여 배치한다.
    col_sort, col_mask, col_save, col_blank = st.columns([2.25, 2.25, 1.45, 4.05])

    with col_sort:
        if st.button("반/번호 순대로 다시 정렬"):
            st.session_state.students = sort_students_df(st.session_state.students)
            st.success("현재 학생 명단을 반/번호 순대로 정렬했습니다.")
            st.rerun()

    with col_mask:
        if st.button("🔒 이름 블라인드 처리", type="primary", help="현재 학생 명단의 성명을 홍*동, 홍*, 홍**동 형식으로 바꿉니다."):
            if st.session_state.students.empty:
                st.warning("블라인드 처리할 학생 명단이 없습니다.")
            else:
                st.session_state.students = mask_student_names_in_df(st.session_state.students, mask_char="*")
                st.warning("학생 이름을 블라인드 처리했습니다. 원래 이름으로 되돌리려면 암호화 전 저장한 JSON 또는 나이스 파일을 다시 불러와야 합니다.")
                st.rerun()

    with col_save:
        if st.button("현재 명단 저장", type="primary"):
            new_df = edited_students.copy()

            for col in ["student_id", "학년", "반", "번호", "성명"]:
                if col not in new_df.columns:
                    new_df[col] = ""

            new_df = new_df[new_df["성명"].astype(str).str.strip() != ""].reset_index(drop=True)

            # 기존 행은 student_id를 유지하고, 새로 추가된 행만 새 ID를 부여한다.
            new_df["student_id"] = new_df["student_id"].apply(
                lambda x: clean_text(x) if clean_text(x) else make_id("stu")
            )

            st.session_state.students = sort_students_df(new_df[["student_id", "학년", "반", "번호", "성명"]])
            st.success("현재 학생 명단을 저장했습니다.")
            st.rerun()

    if not st.session_state.students.empty:
        review_current = st.session_state.students[st.session_state.students["성명"].map(needs_name_review)]
        if not review_current.empty:
            with st.expander(f"현재 명단의 이름 확인 필요 학생 {len(review_current)}명 보기"):
                st.dataframe(
                    review_current.drop(columns=["student_id"], errors="ignore"),
                    use_container_width=True,
                    height=180,
                )


# =========================
# ③ 수행평가 설계
# =========================
with tab3:
    st.subheader("③ 수행평가 설계")

    st.markdown(
        """
        이 화면은 **수행평가**를 먼저 만들고, 각 수행평가 안에 학생별로 입력할 **평가 요소**를 넣는 구조입니다.  
        평가 요소 안에는 필요에 따라 **성취수준 코드, 평가 문구, 개별 코멘트**를 설정합니다.
        """
    )

    with st.expander("➕ 새 수행평가 추가", expanded=True):
        st.caption("먼저 상위 단위인 수행평가를 만들고, 그 안에 평가 요소를 추가합니다.")

        with st.form("add_assessment_form"):
            col1, col2 = st.columns([2, 1])
            with col1:
                new_assessment_name = st.text_input("수행평가명", placeholder="예: 생태지도 만들기")
                new_area = st.text_input("영역/단원", placeholder="예: 생물과 환경")
            with col2:
                new_use = st.checkbox("사용", value=True)
                st.caption("순서는 아래의 드래그 정렬에서 바꿀 수 있습니다.")

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
                            "order": len(st.session_state.assessments) + 1,
                            "use": new_use,
                        }
                    )
                    sanitize_state()
                    st.success("수행평가를 추가했습니다.")
                    st.rerun()

    st.divider()
    st.markdown("### 📚 AI가 참고할 수행평가별 평가 자료")

    if not st.session_state.assessments:
        st.info("아직 등록된 수행평가가 없습니다. 먼저 수행평가를 추가하세요.")

    normalize_assessment_orders()
    sorted_assessments = sorted(
        st.session_state.assessments,
        key=lambda x: int(x.get("order", 999) or 999),
    )

    if len(sorted_assessments) >= 2:
        with st.expander("수행평가 순서 드래그 정렬", expanded=False):
            if sort_items is None:
                st.warning("드래그 정렬 기능을 사용하려면 requirements.txt에 streamlit-sortables를 추가해야 합니다.")
            else:
                st.caption("수행평가명을 마우스로 잡고 위아래로 옮긴 뒤 저장하세요.")
                assessment_labels = [
                    f"{idx}. {assessment.get('name', '이름 없는 수행평가')}"
                    for idx, assessment in enumerate(sorted_assessments, start=1)
                ]
                label_to_assessment_id = {
                    label: assessment.get("assessment_id", "")
                    for label, assessment in zip(assessment_labels, sorted_assessments)
                }

                sorted_labels = sort_labels_with_pretty_box(
                    assessment_labels,
                    key="assessment_drag_sort",
                    header="수행평가 순서",
                )

                if st.button("수행평가 순서 저장"):
                    apply_assessment_drag_order(sorted_labels, label_to_assessment_id)
                    normalize_assessment_orders()
                    st.success("수행평가 순서를 저장했습니다.")
                    st.rerun()

    for assess_index, assessment in enumerate(sorted_assessments, start=1):
        aid = assessment.get("assessment_id", "")
        normalize_item_orders(aid)
        existing_items = sorted(
            get_items_for_assessment(aid),
            key=lambda x: int(x.get("order", 999) or 999),
        )

        status_badge = "사용" if assessment.get("use", True) else "미사용"
        area_text = assessment.get("area", "") or "영역/단원 미입력"
        item_count = len(existing_items)

        with st.container(border=True):
            st.markdown(
                f"""
                ### 📁 수행평가 {assess_index}. {assessment.get('name', '이름 없는 수행평가')}
                **영역/단원:** {area_text} &nbsp;&nbsp;|&nbsp;&nbsp;
                **평가 요소:** {item_count}개 &nbsp;&nbsp;|&nbsp;&nbsp;
                **상태:** {status_badge}
                """
            )

            if assessment.get("description"):
                st.caption(f"활동 설명: {assessment.get('description', '')}")
            else:
                st.caption("활동 설명이 아직 입력되지 않았습니다.")

            with st.expander("⚙️ 수행평가 기본 정보 수정 / 삭제", expanded=False):
                col1, col2, col3 = st.columns([2, 2, 1])

                with col1:
                    assessment["name"] = st.text_input(
                        "수행평가명 수정",
                        value=assessment.get("name", ""),
                        key=f"assess_name_{aid}",
                    )
                    assessment["area"] = st.text_input(
                        "영역/단원 수정",
                        value=assessment.get("area", ""),
                        key=f"assess_area_{aid}",
                    )

                with col2:
                    assessment["description"] = st.text_area(
                        "활동 설명 수정",
                        value=assessment.get("description", ""),
                        key=f"assess_desc_{aid}",
                        height=120,
                    )

                with col3:
                    st.caption("순서는 수행평가 목록 위의 드래그 정렬에서 변경합니다.")
                    assessment["use"] = st.checkbox(
                        "사용",
                        value=assessment.get("use", True),
                        key=f"assess_use_{aid}",
                    )
                    if st.button("수행평가 삭제", key=f"delete_assessment_{aid}"):
                        item_ids = [it.get("item_id", "") for it in get_items_for_assessment(aid)]
                        st.session_state.assessments = [
                            x for x in st.session_state.assessments
                            if x.get("assessment_id", "") != aid
                        ]
                        st.session_state["items"] = [
                            x for x in st.session_state["items"]
                            if x.get("assessment_id", "") != aid
                        ]
                        st.session_state.records = {
                            k: v for k, v in st.session_state.records.items()
                            if k.split("::")[-1] not in item_ids
                        }
                        st.success("수행평가를 삭제했습니다.")
                        st.rerun()

            st.markdown("#### 🧾 평가 요소")

            with st.expander("➕ 이 수행평가에 평가 요소 추가", expanded=(item_count == 0)):
                st.markdown(
                    """
                    수행평가 안에 존재하는 여러 관찰 및 평가 요소들을 추가해주세요.  
                    A, B, C와 같은 **성취도 선택형**으로 개별화시킬 수도 있고, 개인마다 다른 관찰 내용을 적어주는 **개별 코멘트형**으로 더욱 구체적인 개별화가 가능합니다.  
                    또한 이 두 가지를 융합한 **성취도 + 추가 코멘트형**도 가능합니다.
                    """
                )

                item_name = st.text_input(
                    "평가 요소명",
                    placeholder="예: 생태지도 결과물 평가",
                    key=f"new_item_name_{aid}",
                )

                item_type_label = st.selectbox(
                    "기록 방식",
                    ["성취도 선택형", "개별 코멘트형", "성취도 + 추가 코멘트형"],
                    key=f"new_item_type_{aid}",
                    help="개별 코멘트형을 선택하면 성취수준 코드와 루브릭 입력칸이 사라집니다.",
                )
                item_type = item_type_from_kor(item_type_label)

                levels = []
                rubrics = {}

                if item_type != "comment":
                    levels, rubrics = render_rubric_input_block(
                        prefix=f"new_item_rubric_{aid}",
                        current_levels=["A", "B", "C", "D", "E"],
                        current_rubrics={
                            "A": "우수한 수준으로 수행함",
                            "B": "대체로 적절하게 수행함",
                            "C": "일부 보완이 필요함",
                            "D": "기본적인 참여가 이루어짐",
                            "E": "지속적인 보완이 필요함",
                        },
                    )
                else:
                    st.info("개별 코멘트형입니다. 성취수준 코드 없이 학생별 서술형 코멘트만 입력합니다.")

                item_order = len(get_items_for_assessment(aid)) + 1

                if st.button("이 수행평가에 평가 요소 추가", key=f"add_item_button_{aid}"):
                    if not item_name.strip():
                        st.warning("평가 요소명을 입력하세요.")
                    else:
                        if item_type == "comment":
                            levels, rubrics = [], {}

                        st.session_state["items"].append(
                            {
                                "item_id": make_id("item"),
                                "assessment_id": aid,
                                "name": item_name.strip(),
                                "type": item_type,
                                "levels": levels,
                                "rubrics": rubrics,
                                "order": int(item_order),
                            }
                        )
                        normalize_item_orders(aid)
                        sanitize_state()
                        st.success("평가 요소를 추가했습니다.")
                        st.rerun()

            if existing_items:
                if len(existing_items) >= 2:
                    with st.expander("평가 요소 순서 드래그 정렬", expanded=True):
                        if sort_items is None:
                            st.warning("드래그 정렬 기능을 사용하려면 requirements.txt에 streamlit-sortables를 추가해야 합니다.")
                        else:
                            st.caption("평가 요소명을 마우스로 잡고 위아래로 옮긴 뒤 저장하세요.")
                            item_labels = [
                                f"{idx}. {item.get('name', '이름 없는 평가 요소')}"
                                for idx, item in enumerate(existing_items, start=1)
                            ]
                            label_to_item_id = {
                                label: item.get("item_id", "")
                                for label, item in zip(item_labels, existing_items)
                            }

                            sorted_item_labels = sort_labels_with_pretty_box(
                                item_labels,
                                key=f"item_drag_sort_{aid}",
                                header="평가 요소 순서",
                            )

                            if st.button("평가 요소 순서 저장", key=f"save_item_drag_sort_{aid}"):
                                apply_item_drag_order(aid, sorted_item_labels, label_to_item_id)
                                normalize_item_orders(aid)
                                st.success("평가 요소 순서를 저장했습니다.")
                                st.rerun()

                for item_index, item in enumerate(existing_items, start=1):
                    item_id = item.get("item_id", "")
                    item_type_label = item_type_to_kor(item.get("type", "rubric"))
                    level_count = len(item.get("levels", [])) if item.get("type") != "comment" else 0

                    with st.container(border=True):
                        st.markdown(
                            f"##### 🧾 평가 요소 {item_index}. {item.get('name', '이름 없는 평가 요소')}"
                        )

                        col1, col2, col3 = st.columns([2.2, 1.6, 1])

                        with col1:
                            item["name"] = st.text_input(
                                "항목명",
                                value=item.get("name", ""),
                                key=f"item_name_{item_id}",
                            )

                        with col2:
                            current_type_label = item_type_to_kor(item.get("type", "rubric"))
                            type_options = ["성취도 선택형", "개별 코멘트형", "성취도 + 추가 코멘트형"]
                            new_type_label = st.selectbox(
                                "기록 방식",
                                type_options,
                                index=type_options.index(current_type_label) if current_type_label in type_options else 0,
                                key=f"item_type_{item_id}",
                            )
                            new_type = item_type_from_kor(new_type_label)

                            if new_type != item.get("type"):
                                item["type"] = new_type
                                if new_type == "comment":
                                    item["levels"] = []
                                    item["rubrics"] = {}
                                elif not item.get("levels"):
                                    item["levels"] = ["A", "B", "C", "D", "E"]
                                    item["rubrics"] = {level: "" for level in item["levels"]}

                        with col3:
                            st.caption(f"현재 {item_index}번째")
                            if st.button("평가 요소 삭제", key=f"delete_item_{item_id}"):
                                st.session_state["items"] = [
                                    x for x in st.session_state["items"]
                                    if x.get("item_id", "") != item_id
                                ]
                                st.session_state.records = {
                                    k: v for k, v in st.session_state.records.items()
                                    if not k.endswith(f"::{item_id}")
                                }
                                normalize_item_orders(aid)
                                st.success("평가 요소를 삭제했습니다.")
                                st.rerun()

                        if item.get("type") in ["rubric", "rubric_plus"]:
                            levels, rubrics = render_rubric_input_block(
                                prefix=f"edit_item_rubric_{item_id}",
                                current_levels=item.get("levels", []),
                                current_rubrics=item.get("rubrics", {}),
                            )

                            if st.button("성취수준/평가 문구 저장", key=f"save_rubric_{item_id}"):
                                item["levels"] = levels
                                item["rubrics"] = rubrics
                                st.success("성취수준과 평가 문구를 저장했습니다.")
                                st.rerun()
                        else:
                            st.info("개별 코멘트형 항목입니다. 학생별 기록 입력 화면에서 학생별 서술형 코멘트를 입력합니다.")
            else:
                st.info("아직 이 수행평가에 등록된 평가 요소가 없습니다. 위의 '평가 요소 추가'를 눌러 항목을 추가하세요.")



# =========================
# ④ 학생별 기록 입력
# =========================
with tab4:
    st.subheader("④ 학생별 기록 입력")

    if st.session_state.students.empty:
        st.warning("먼저 학생 명단을 업로드하거나 입력하세요.")

    elif not st.session_state["items"]:
        st.warning("먼저 수행평가와 평가 요소를 추가하세요.")

    else:
        students = st.session_state.students.copy()
        assessments = [a for a in st.session_state.assessments if a.get("use", True)]
        assess_options = ["전체"] + [a.get("name", "") for a in assessments]
        selected_assess_name = st.selectbox("수행평가 필터", assess_options)

        if selected_assess_name == "전체":
            selected_assess_ids = [a.get("assessment_id", "") for a in assessments]
        else:
            selected_assess_ids = [
                a.get("assessment_id", "")
                for a in assessments
                if a.get("name", "") == selected_assess_name
            ]

        selected_items = [
            item for item in st.session_state["items"]
            if isinstance(item, dict) and item.get("assessment_id", "") in selected_assess_ids
        ]
        selected_items = sorted(
            selected_items,
            key=lambda x: (get_assessment_name(x.get("assessment_id", "")), int(x.get("order", 999) or 999)),
        )

        if not selected_items:
            st.info("선택한 수행평가에 평가 요소가 없습니다.")
        else:
            data_rows = []
            for _, student in students.iterrows():
                row = {
                    "student_id": student["student_id"],
                    "학년": student.get("학년", ""),
                    "반": student.get("반", ""),
                    "번호": student.get("번호", ""),
                    "성명": student.get("성명", ""),
                }

                for item in selected_items:
                    rec = get_record(student["student_id"], item.get("item_id", ""))
                    base_col = f"{get_assessment_name(item.get('assessment_id', ''))} - {item.get('name', '')}"

                    if item.get("type") == "rubric":
                        row[base_col] = rec.get("level", "")
                    elif item.get("type") == "comment":
                        row[base_col] = rec.get("comment", "")
                    elif item.get("type") == "rubric_plus":
                        row[f"{base_col} 성취수준"] = rec.get("level", "")
                        row[f"{base_col} 추가코멘트"] = rec.get("comment", "")

                data_rows.append(row)

            input_df = pd.DataFrame(data_rows)

            column_config = {}
            for item in selected_items:
                base_col = f"{get_assessment_name(item.get('assessment_id', ''))} - {item.get('name', '')}"

                if item.get("type") == "rubric":
                    column_config[base_col] = st.column_config.SelectboxColumn(
                        base_col,
                        options=[""] + item.get("levels", []),
                        required=False,
                    )
                elif item.get("type") == "rubric_plus":
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
                        base_col = f"{get_assessment_name(item.get('assessment_id', ''))} - {item.get('name', '')}"

                        if item.get("type") == "rubric":
                            set_record(sid, item.get("item_id", ""), level=row.get(base_col, ""), comment="")
                        elif item.get("type") == "comment":
                            set_record(sid, item.get("item_id", ""), level="", comment=row.get(base_col, ""))
                        elif item.get("type") == "rubric_plus":
                            set_record(
                                sid,
                                item.get("item_id", ""),
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

        col1, col2 = st.columns([2, 1])
        with col1:
            api_key = st.text_input(
                "OpenAI API Key",
                value=default_api_key,
                type="password",
                help="입력하지 않으면 API 없이 간단 조합 방식으로 생성됩니다.",
            )
        with col2:
            model = st.text_input("모델명", value="gpt-4o-mini")

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
            job = st.session_state.generation_job

            if not job.get("active", False):
                if st.button("전체 학생 생기부 생성 시작"):
                    st.session_state.generation_job = {
                        "active": True,
                        "stop_requested": False,
                        "student_ids": students["student_id"].tolist(),
                        "index": 0,
                        "log": [],
                        "api_key": api_key,
                        "model": model,
                        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "finished_at": "",
                    }
                    st.success("전체 생성 작업을 시작합니다.")
                    st.rerun()
            else:
                if st.button("생기부 생성 중지", type="secondary"):
                    job["active"] = False
                    job["stop_requested"] = True
                    job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    st.session_state.generation_job = job
                    st.warning("생성 중지를 요청했습니다. 이미 생성된 학생 결과는 저장되어 있습니다.")
                    st.rerun()

        # 전체 생성 작업 상태 표시 및 한 명씩 자동 처리
        job = st.session_state.generation_job
        if job.get("active", False) or job.get("log"):
            st.markdown("#### 전체 생성 진행 상황")
            total = len(job.get("student_ids", []))
            done = int(job.get("index", 0))

            if total > 0:
                st.progress(min(done / total, 1.0))
                st.caption(f"진행률: {done}/{total}명")

            if job.get("active", False):
                st.info("전체 생성이 진행 중입니다. 중지하려면 오른쪽의 '생기부 생성 중지' 버튼을 누르세요. 현재 처리 중인 학생은 완료된 뒤 멈춥니다.")
            elif job.get("stop_requested", False):
                st.warning("전체 생성이 중지되었습니다. 중지 전까지 생성된 문구는 아래 결과표와 다운로드 엑셀에 자동 반영됩니다.")
            elif job.get("log"):
                st.success("전체 생성 작업이 완료되었습니다.")

            if job.get("log"):
                st.dataframe(
                    pd.DataFrame(job.get("log", [])),
                    use_container_width=True,
                    height=360,
                )

        if job.get("active", False) and not job.get("stop_requested", False):
            student_ids = job.get("student_ids", [])
            index = int(job.get("index", 0))
            total = len(student_ids)

            if index >= total:
                job["active"] = False
                job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.session_state.generation_job = job
                st.success("전체 학생 생기부 생성을 완료했습니다.")
                st.rerun()

            current_sid = student_ids[index]
            matched = students[students["student_id"] == current_sid]

            if matched.empty:
                job["index"] = index + 1
                st.session_state.generation_job = job
                st.rerun()

            student = matched.iloc[0]
            label = f"{student.get('반', '')}반 {student.get('번호', '')}번 {student.get('성명', '')}"

            with st.spinner(f"{index + 1}/{total} 생성 중: {label}"):
                material = build_student_material(student)
                prompt = build_prompt(material)

                generated = None
                if job.get("api_key"):
                    generated = generate_with_openai(prompt, job.get("api_key", ""), job.get("model", model))

                if not generated:
                    generated = fallback_generate(material)

            st.session_state.results[current_sid] = {
                "material": material,
                "generated": generated,
                "edited": generated,
                "bytes": byte_count(generated),
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            job.setdefault("log", []).append(
                {
                    "순서": index + 1,
                    "반": student.get("반", ""),
                    "번호": student.get("번호", ""),
                    "성명": student.get("성명", ""),
                    "생성 문구": generated,
                    "byte": byte_count(generated),
                    "상태": "저장 완료",
                }
            )
            job["index"] = index + 1

            if job["index"] >= total:
                job["active"] = False
                job["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.session_state.generation_job = job
                st.success("전체 학생 생기부 생성을 완료했습니다.")
            else:
                st.session_state.generation_job = job
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
            current_bytes = byte_count(edited)
            st.caption(f"현재 byte: {current_bytes}")

            if current_bytes < st.session_state.settings["target_bytes_min"]:
                st.warning("목표 byte보다 짧습니다.")
            elif current_bytes > st.session_state.settings["target_bytes_max"]:
                st.warning("목표 byte보다 깁니다.")
            else:
                st.success("목표 byte 범위 안에 있습니다.")

            if st.button("수정 문구 저장"):
                result["edited"] = edited
                result["bytes"] = current_bytes
                st.session_state.results[sid] = result
                st.success("수정 문구를 저장했습니다.")
        else:
            st.info("아직 이 학생의 생기부 문구가 생성되지 않았습니다.")

        st.divider()
        st.markdown("#### 전체 결과표")

        result_rows = []
        for _, student in students.iterrows():
            result = st.session_state.results.get(student["student_id"], {})
            text = result.get("edited", result.get("generated", ""))
            result_rows.append(
                {
                    "학년": student.get("학년", ""),
                    "반": student.get("반", ""),
                    "번호": student.get("번호", ""),
                    "성명": student.get("성명", ""),
                    "생성/수정 문구": text,
                    "byte": byte_count(text),
                    "생성일시": result.get("created_at", ""),
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
