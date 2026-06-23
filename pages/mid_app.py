import html
import json
import re
import uuid
from datetime import datetime
from io import BytesIO
from urllib import error as urllib_error
from urllib import parse as urllib_parse
from urllib import request as urllib_request

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =========================
# 중학교 간편 생기부 v07
# =========================
st.set_page_config(
    page_title="중학교 간편 생기부",
    page_icon="🍯",
    layout="wide",
)

MID_APP_TITLE = "🍯 중학교 간편 생기부 v07"
MID_APP_SUBTITLE = "수행평가·관찰 영역 기반 중학교 생기부 간편 작성 도우미 · patched-20260623-mid-v07"
MID_APP_VERSION = "patched-20260623-mid-v07"

MID_DEFAULT_RULES = """- 중학교 학교생활기록부 교과 세부능력 및 특기사항 문체로 작성한다.
- 학생 이름, 학년, 반, 번호, 학교명 등 개인정보를 쓰지 않는다.
- 첫 문장을 '학생은', '이 학생은', '해당 학생은'으로 시작하지 않는다.
- 활동 결과에 없는 내용을 추측하거나 과장하지 않는다.
- '깊은 이해', '창의융합', '혁신적', '흥미와 전문성 심화', '본인은', '의지를 밝힘' 같은 표현을 피한다.
- 한 문장 또는 짧은 한 문단으로 작성한다.
- 명사형 종결을 사용한다. 예: 수행함, 설명함, 정리함, 제시함, 해석함, 이해한 것으로 보임."""

MASTER_PROMPT = """너는 중학교 학교생활기록부 교과 세부능력 및 특기사항을 작성하는 교사 보조 도구이다.

[기본 원칙]
- 제공된 수행평가와 관찰 영역별 성취수준 자료만 근거로 사용한다.
- 학생 이름, 학년, 반, 번호, 학교명 등 개인정보는 포함하지 않는다.
- 학생은, 이 학생은, 해당 학생은으로 문장을 시작하지 않는다.
- 성취수준 코드를 그대로 나열하지 않고 교사의 평가 문구를 자연스럽게 바꾸어 작성한다.
- 근거 없는 인성 평가, 성격 판단, 진로 추정, 태도 평가는 작성하지 않는다.
- 부정적 표현을 직접 쓰지 않고 현재 수행한 내용 중심으로 서술한다.
- 한 문단으로 작성한다.
- 문장은 명사형 종결 어미로 마무리한다.
- 제목, 번호, 설명, 따옴표, 안내 문구 없이 생기부 문장만 출력한다."""

AI_PROVIDER_OPTIONS = ["ChatGPT", "Gemini", "Claude"]
AI_DEFAULT_MODELS = {
    "ChatGPT": "gpt-5.5",
    "Gemini": "gemini-3.1-pro-preview",
    "Claude": "claude-fable-5",
}
AI_SECRET_KEY_NAMES = {
    "ChatGPT": "OPENAI_API_KEY",
    "Gemini": "GEMINI_API_KEY",
    "Claude": "ANTHROPIC_API_KEY",
}

VARIATION_GUIDES = {
    "낮음": "평가 내용은 그대로 유지하고 어휘와 문장 순서만 조금씩 바꾼다.",
    "보통": "평가 내용은 유지하되 문장 구조와 표현을 적당히 다르게 만든다.",
    "높음": "평가 내용은 벗어나지 않으면서 문장 흐름과 표현을 다양하게 구성한다.",
}


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


def json_safe(obj):
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


def default_level_code(index: int) -> str:
    default_codes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    if 0 <= index < len(default_codes):
        return default_codes[index]
    return str(index + 1)


def as_dict(value):
    """이전 버전 세션/JSON에서 list 등으로 남은 값을 안전하게 dict로 바꾼다."""
    return value if isinstance(value, dict) else {}


def get_default_ai_model(provider: str) -> str:
    provider = provider if provider in AI_DEFAULT_MODELS else "ChatGPT"
    return AI_DEFAULT_MODELS[provider]


def get_default_ai_key(provider: str) -> str:
    secret_name = AI_SECRET_KEY_NAMES.get(provider, "OPENAI_API_KEY")
    try:
        return st.secrets.get(secret_name, "")
    except Exception:
        return ""


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


def sort_assessments():
    return sorted(
        st.session_state.get("mid_assessments", []),
        key=lambda x: (
            int(x.get("order", 999) or 999),
            clean_text(x.get("name", "")),
            clean_text(x.get("assessment_id", "")),
        ),
    )


def get_items_for_assessment(assessment_id):
    return [
        item for item in st.session_state.get("mid_items", [])
        if isinstance(item, dict) and item.get("assessment_id", "") == assessment_id
    ]


def sort_items_for_assessment(assessment_id):
    return sorted(
        get_items_for_assessment(assessment_id),
        key=lambda x: (
            int(x.get("order", 999) or 999),
            clean_text(x.get("name", "")),
            clean_text(x.get("item_id", "")),
        ),
    )


def get_assessment_name(assessment_id):
    for assessment in st.session_state.get("mid_assessments", []):
        if assessment.get("assessment_id", "") == assessment_id:
            return assessment.get("name", "")
    return ""


def get_item_by_id(item_id):
    for item in st.session_state.get("mid_items", []):
        if item.get("item_id", "") == item_id:
            return item
    return None


def all_used_items():
    rows = []
    for assessment in sort_assessments():
        if not assessment.get("use", True):
            continue
        for item in sort_items_for_assessment(assessment.get("assessment_id", "")):
            rows.append(item)
    return rows


def record_key(student_id, item_id):
    return f"{student_id}::{item_id}"


def get_level(student_id, item_id):
    return clean_text(st.session_state.get("mid_records", {}).get(record_key(student_id, item_id), ""))


def set_level(student_id, item_id, level):
    st.session_state.mid_records[record_key(student_id, item_id)] = clean_text(level)


def normalize_sentence(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"^[0-9]+[\.)]\s*", "", text)
    return text.strip(' \n\t-•·"\'“”‘’')


# =========================
# 세션 상태
# =========================
def init_mid_state():
    if "mid_settings" not in st.session_state:
        st.session_state.mid_settings = {
            "school_year": "2026",
            "semester": "1학기",
            "school_level": "중학교",
            "grade": "2",
            "subject": "과학",
            "target_bytes_min": 250,
            "target_bytes_max": 450,
            "custom_rules": MID_DEFAULT_RULES,
        }

    if "mid_assessments" not in st.session_state:
        st.session_state.mid_assessments = []

    if "mid_items" not in st.session_state:
        st.session_state.mid_items = []

    if "mid_students" not in st.session_state:
        st.session_state.mid_students = pd.DataFrame(
            columns=["student_id", "학년", "반", "번호", "성명"]
        )

    if "mid_records" not in st.session_state:
        st.session_state.mid_records = {}

    if "mid_results" not in st.session_state:
        st.session_state.mid_results = {}

    if "mid_selected_result_student_id" not in st.session_state:
        st.session_state.mid_selected_result_student_id = ""


def sanitize_mid_state():
    settings = st.session_state.get("mid_settings", {})
    if "custom_rules" not in settings:
        settings["custom_rules"] = MID_DEFAULT_RULES
    if "target_bytes_min" not in settings:
        settings["target_bytes_min"] = 250
    if "target_bytes_max" not in settings:
        settings["target_bytes_max"] = 450
    st.session_state.mid_settings = settings

    if not isinstance(st.session_state.get("mid_students"), pd.DataFrame):
        st.session_state.mid_students = pd.DataFrame(st.session_state.get("mid_students", []))
    for col in ["student_id", "학년", "반", "번호", "성명"]:
        if col not in st.session_state.mid_students.columns:
            st.session_state.mid_students[col] = ""
    if not st.session_state.mid_students.empty:
        st.session_state.mid_students["student_id"] = st.session_state.mid_students["student_id"].apply(
            lambda x: clean_text(x) if clean_text(x) else make_id("mid_stu")
        )
        st.session_state.mid_students = sort_students_df(st.session_state.mid_students)

    clean_assessments = []
    for idx, assessment in enumerate(st.session_state.get("mid_assessments", []), start=1):
        if not isinstance(assessment, dict):
            continue
        if not clean_text(assessment.get("assessment_id", "")):
            assessment["assessment_id"] = make_id("mid_assess")
        assessment["name"] = clean_text(assessment.get("name", "")) or "이름 없는 수행평가"
        assessment["unit"] = clean_text(assessment.get("unit", ""))
        assessment["description"] = clean_text(assessment.get("description", ""))
        assessment["order"] = int(assessment.get("order", idx) or idx)
        assessment["use"] = bool(assessment.get("use", True))
        clean_assessments.append(assessment)
    clean_assessments = sorted(clean_assessments, key=lambda x: int(x.get("order", 999) or 999))
    for idx, assessment in enumerate(clean_assessments, start=1):
        assessment["order"] = idx
    st.session_state.mid_assessments = clean_assessments
    valid_assessment_ids = {a["assessment_id"] for a in clean_assessments}

    clean_items = []
    for idx, item in enumerate(st.session_state.get("mid_items", []), start=1):
        if not isinstance(item, dict):
            continue
        if item.get("assessment_id", "") not in valid_assessment_ids:
            continue
        if not clean_text(item.get("item_id", "")):
            item["item_id"] = make_id("mid_item")
        item["name"] = clean_text(item.get("name", "")) or "이름 없는 관찰 영역"
        item["order"] = int(item.get("order", idx) or idx)
        if not isinstance(item.get("levels", []), list):
            item["levels"] = []
        if not isinstance(item.get("rubrics", {}), dict):
            item["rubrics"] = {}
        clean_items.append(item)
    st.session_state.mid_items = clean_items
    for assessment in st.session_state.mid_assessments:
        aid = assessment.get("assessment_id", "")
        items = sort_items_for_assessment(aid)
        for idx, item in enumerate(items, start=1):
            item["order"] = idx

    valid_student_ids = set(st.session_state.mid_students["student_id"].astype(str).tolist()) if not st.session_state.mid_students.empty else set()
    valid_item_ids = {item.get("item_id", "") for item in st.session_state.mid_items}

    raw_records = as_dict(st.session_state.get("mid_records", {}))
    clean_records = {}
    for key, value in raw_records.items():
        key = str(key)
        if "::" not in key:
            continue
        sid, item_id = key.split("::", 1)
        if sid in valid_student_ids and item_id in valid_item_ids:
            clean_records[key] = clean_text(value)
    st.session_state.mid_records = clean_records

    raw_results = st.session_state.get("mid_results", {})
    # v01~v02에서는 mid_results가 list였기 때문에 v03 진입 시 AttributeError가 날 수 있다.
    # v04부터는 dict가 아닐 경우 빈 결과로 안전하게 초기화한다.
    raw_results = raw_results if isinstance(raw_results, dict) else {}
    clean_results = {}
    for sid, result in raw_results.items():
        sid = clean_text(sid)
        if sid in valid_student_ids and isinstance(result, dict):
            clean_results[sid] = result
    st.session_state.mid_results = clean_results


init_mid_state()
sanitize_mid_state()


# =========================
# 프로젝트 저장/불러오기
# =========================
def mid_project_to_json() -> str:
    data = {
        "settings": st.session_state.mid_settings,
        "students": st.session_state.mid_students.to_dict(orient="records"),
        "assessments": st.session_state.mid_assessments,
        "items": st.session_state.mid_items,
        "records": st.session_state.mid_records,
        "results": st.session_state.mid_results,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "개꿀 생기부 - 중학교 간편",
        "version": MID_APP_VERSION,
    }
    return json.dumps(json_safe(data), ensure_ascii=False, indent=2, default=str)


def apply_mid_project_data(data):
    st.session_state.mid_settings = data.get("settings", st.session_state.mid_settings)
    st.session_state.mid_students = pd.DataFrame(data.get("students", []))
    st.session_state.mid_assessments = data.get("assessments", [])
    st.session_state.mid_items = data.get("items", [])
    st.session_state.mid_records = data.get("records", {})
    st.session_state.mid_results = data.get("results", {})
    st.session_state.mid_selected_result_student_id = ""
    sanitize_mid_state()


def load_mid_project_json(uploaded_file):
    data = json.load(uploaded_file)
    apply_mid_project_data(data)


def build_mid_sample_project_data():
    assessments = [
        {
            "assessment_id": "mid_assess_digest",
            "name": "소화 기관 모형 만들기",
            "unit": "동물의 몸과 영양소",
            "description": "소화 기관의 구조와 각 기관의 역할을 모형으로 표현하고 설명하는 활동",
            "order": 1,
            "use": True,
        },
        {
            "assessment_id": "mid_assess_photo",
            "name": "광합성 조건 탐구하기",
            "unit": "식물과 에너지",
            "description": "광합성에 영향을 주는 환경 요인을 정하고 변인을 통제하여 실험을 설계하는 활동",
            "order": 2,
            "use": True,
        },
    ]
    common_levels = ["A", "B", "C"]
    common_rubrics = {
        "A": "핵심 개념을 정확히 연결하여 활동 결과를 구체적으로 설명함",
        "B": "주요 개념을 바탕으로 활동 결과를 대체로 적절하게 설명함",
        "C": "기초 개념을 중심으로 활동에 참여하고 기본 내용을 정리함",
    }
    items = [
        {
            "item_id": "mid_item_digest_order",
            "assessment_id": "mid_assess_digest",
            "name": "소화 기관의 순서와 역할",
            "levels": common_levels,
            "rubrics": common_rubrics,
            "order": 1,
        },
        {
            "item_id": "mid_item_digest_model",
            "assessment_id": "mid_assess_digest",
            "name": "모형 표현과 설명",
            "levels": common_levels,
            "rubrics": common_rubrics,
            "order": 2,
        },
        {
            "item_id": "mid_item_photo_variable",
            "assessment_id": "mid_assess_photo",
            "name": "변인 통제",
            "levels": common_levels,
            "rubrics": common_rubrics,
            "order": 1,
        },
        {
            "item_id": "mid_item_photo_predict",
            "assessment_id": "mid_assess_photo",
            "name": "결과 예측과 해석",
            "levels": common_levels,
            "rubrics": common_rubrics,
            "order": 2,
        },
    ]
    names = ["김민준", "박서연", "최도윤", "이하은", "정우진", "한지우", "오서준", "윤채아", "임하준", "강민서"]
    students = []
    records = {}
    level_cycle = ["A", "B", "A", "B", "C", "B", "A", "C", "B", "A"]
    for idx, name in enumerate(names, start=1):
        sid = f"mid_sample_stu_{idx:02d}"
        students.append({"student_id": sid, "학년": "2", "반": "1", "번호": str(idx), "성명": name})
        for j, item in enumerate(items):
            records[f"{sid}::{item['item_id']}"] = level_cycle[(idx + j - 1) % len(level_cycle)]
    return {
        "settings": {
            "school_year": "2026",
            "semester": "1학기",
            "school_level": "중학교",
            "grade": "2",
            "subject": "과학",
            "target_bytes_min": 250,
            "target_bytes_max": 450,
            "custom_rules": MID_DEFAULT_RULES,
        },
        "students": students,
        "assessments": assessments,
        "items": items,
        "records": records,
        "results": {},
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "개꿀 생기부 - 중학교 간편",
        "version": "sample-mid-v05",
    }


def load_mid_sample_data():
    apply_mid_project_data(build_mid_sample_project_data())


# =========================
# API 생성
# =========================
def _post_json(url, headers, payload, timeout=90):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib_request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib_request.urlopen(req, timeout=timeout) as response:
            body = response.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib_error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {e.code}: {detail[:700]}") from e


def generate_with_openai(prompt, api_key, model):
    if not api_key:
        return None
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        response = client.responses.create(model=model, input=prompt)
        return response.output_text.strip()
    except Exception as e:
        st.error(f"ChatGPT API 생성 중 오류가 발생했습니다: {e}")
        return None


def generate_with_gemini(prompt, api_key, model):
    if not api_key:
        return None
    try:
        model_path = urllib_parse.quote(clean_text(model), safe="")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_path}:generateContent?key={urllib_parse.quote(api_key)}"
        payload = {
            "contents": [{"role": "user", "parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.35},
        }
        data = _post_json(url, headers={"Content-Type": "application/json"}, payload=payload)
        parts = data.get("candidates", [{}])[0].get("content", {}).get("parts", [])
        text = "".join([part.get("text", "") for part in parts if isinstance(part, dict)]).strip()
        return text or None
    except Exception as e:
        st.error(f"Gemini API 생성 중 오류가 발생했습니다: {e}")
        return None


def generate_with_claude(prompt, api_key, model):
    if not api_key:
        return None
    try:
        payload = {
            "model": model,
            "max_tokens": 2000,
            "temperature": 0.35,
            "messages": [{"role": "user", "content": prompt}],
        }
        data = _post_json(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            payload=payload,
        )
        blocks = data.get("content", [])
        text = "".join([
            block.get("text", "")
            for block in blocks
            if isinstance(block, dict) and block.get("type") == "text"
        ]).strip()
        return text or None
    except Exception as e:
        st.error(f"Claude API 생성 중 오류가 발생했습니다: {e}")
        return None


def generate_with_ai(prompt, ai_provider, api_key, model):
    provider = ai_provider if ai_provider in AI_PROVIDER_OPTIONS else "ChatGPT"
    model = clean_text(model) or get_default_ai_model(provider)
    if provider == "Gemini":
        return generate_with_gemini(prompt, api_key, model)
    if provider == "Claude":
        return generate_with_claude(prompt, api_key, model)
    return generate_with_openai(prompt, api_key, model)


def build_student_material(student):
    sid = student.get("student_id", "")
    lines = []
    # 개인정보 보호를 위해 이름/학년/반/번호는 프롬프트에 넣지 않는다.
    for assessment in sort_assessments():
        if not assessment.get("use", True):
            continue
        chunks = []
        for item in sort_items_for_assessment(assessment.get("assessment_id", "")):
            level = get_level(sid, item.get("item_id", ""))
            rubrics = item.get("rubrics", {}) if isinstance(item.get("rubrics", {}), dict) else {}
            teacher_comment = clean_text(rubrics.get(level, ""))
            levels = [clean_text(x) for x in item.get("levels", []) if clean_text(x)]
            if level or teacher_comment:
                if levels:
                    level_text = f"전체 성취수준 {', '.join(levels)} 중 {level or '미입력'}"
                else:
                    level_text = f"성취수준 {level or '미입력'}"
                chunks.append(f"- {item.get('name', '')}: {level_text} / 교사의 평가: {teacher_comment}")
        if chunks:
            lines.append(f"{assessment.get('name', '')}")
            if assessment.get("unit"):
                lines.append(f"- 단원: {assessment.get('unit', '')}")
            if assessment.get("description"):
                lines.append(f"- 활동 설명: {assessment.get('description', '')}")
            lines.extend(chunks)
            lines.append("")
    return "\n".join(lines).strip()


def build_prompt(material, variation_level="보통", variant_no=1):
    settings = st.session_state.mid_settings
    custom_rules = clean_text(settings.get("custom_rules", MID_DEFAULT_RULES))
    variation_guide = VARIATION_GUIDES.get(variation_level, VARIATION_GUIDES["보통"])
    return f"""
{MASTER_PROMPT}

[선생님 추가 작성 규칙]
{custom_rules}

[작성 조건]
- 목표 분량: {settings.get('target_bytes_min', 250)}~{settings.get('target_bytes_max', 450)} byte
- 중학교 생기부에 맞게 간결하게 작성한다.
- 같은 성취수준 학생끼리도 결과 문장이 완전히 같지 않도록 자연스럽게 변주한다.
- 표현 변주 기준: {variation_guide}
- 문장 변주 번호: {variant_no}

[평가 자료]
{material}

[최종 출력]
세부능력 및 특기사항 문장만 출력하라.
""".strip()


def fallback_generate(material, variant_no=1):
    pieces = []
    for line in material.splitlines():
        line = line.strip()
        if not line.startswith("- "):
            continue
        if "교사의 평가:" in line:
            value = line.split("교사의 평가:", 1)[1].strip()
            if value:
                pieces.append(value.rstrip("."))
    unique = []
    seen = set()
    for piece in pieces:
        if piece and piece not in seen:
            unique.append(piece)
            seen.add(piece)
    if not unique:
        return "입력된 성취수준 자료가 부족하여 생기부 문장 생성이 어려움."

    if variant_no % 3 == 1:
        text = ", ".join(unique[:4])
        if not text.endswith(("함", "임", "음", "보임")):
            text += "함"
        return text
    if variant_no % 3 == 2:
        text = " 및 ".join(unique[:3])
        if not text.endswith(("함", "임", "음", "보임")):
            text += "을 바탕으로 활동을 수행함"
        return text
    text = ". ".join(unique[:3])
    if not text.endswith(("함", "임", "음", "보임", ".")):
        text += "함"
    return text


def generate_for_student(student, ai_provider, api_key, model, variation_level, variant_no):
    material = build_student_material(student)
    prompt = build_prompt(material, variation_level=variation_level, variant_no=variant_no)
    generated = None
    if api_key:
        generated = generate_with_ai(prompt, ai_provider, api_key, model)
    if not generated:
        generated = fallback_generate(material, variant_no=variant_no)
    generated = normalize_sentence(generated)
    return material, generated


# =========================
# 입력표 / 엑셀
# =========================
def build_record_matrix_df():
    students = st.session_state.mid_students.copy()
    items = all_used_items()
    if students.empty:
        students = pd.DataFrame([
            {
                "student_id": make_id("mid_stu"),
                "학년": st.session_state.mid_settings.get("grade", "2"),
                "반": "1",
                "번호": "1",
                "성명": "",
            }
        ])

    rows = []
    for _, student in students.iterrows():
        sid = clean_text(student.get("student_id", "")) or make_id("mid_stu")
        row = {
            "student_id": sid,
            "학년": clean_text(student.get("학년", "")),
            "반": clean_text(student.get("반", "")),
            "번호": clean_text(student.get("번호", "")),
            "성명": clean_text(student.get("성명", "")),
        }
        for item in items:
            row[item.get("item_id", "")] = get_level(sid, item.get("item_id", ""))
        rows.append(row)
    return pd.DataFrame(rows), items


def save_record_matrix_df(edited_df, items):
    if edited_df.empty:
        return 0
    students = edited_df[["student_id", "학년", "반", "번호", "성명"]].copy()
    students = students[students["성명"].astype(str).str.strip() != ""].reset_index(drop=True)
    students["student_id"] = students["student_id"].apply(lambda x: clean_text(x) if clean_text(x) else make_id("mid_stu"))
    st.session_state.mid_students = sort_students_df(students)

    valid_student_ids = set(st.session_state.mid_students["student_id"].astype(str).tolist())
    valid_item_ids = {item.get("item_id", "") for item in items}
    new_records = {}
    saved_count = 0
    for _, row in edited_df.iterrows():
        sid = clean_text(row.get("student_id", ""))
        if sid not in valid_student_ids:
            continue
        for item in items:
            item_id = item.get("item_id", "")
            if item_id not in valid_item_ids:
                continue
            value = clean_text(row.get(item_id, ""))
            if value:
                new_records[record_key(sid, item_id)] = value
            saved_count += 1
    st.session_state.mid_records = new_records
    sanitize_mid_state()
    return saved_count


def make_record_input_excel():
    students = st.session_state.mid_students.copy()
    items = all_used_items()
    wb = Workbook()
    ws = wb.active
    ws.title = "학생별성취수준"
    list_ws = wb.create_sheet("선택목록")
    list_ws.sheet_state = "hidden"

    base_headers = ["student_id", "학년", "반", "번호", "성명"]
    start_row = 4
    data_start_row = 6
    item_start_col = 6
    last_col = max(item_start_col + len(items) - 1, 5)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.cell(1, 1).value = "중학교 생기부 학생별 성취수준 입력표"
    ws.cell(1, 1).font = Font(bold=True, size=14, color="111827")
    ws.cell(1, 1).alignment = Alignment(vertical="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws.cell(2, 1).value = "윗줄은 수행평가명, 아랫줄은 관찰 영역명입니다. 성취수준 칸은 웹앱에서 설정한 코드 기준으로 입력하세요."
    ws.cell(2, 1).font = Font(size=10, color="6B7280")
    ws.cell(2, 1).alignment = Alignment(wrap_text=True)

    for col_idx, header in enumerate(base_headers, start=1):
        ws.cell(start_row, col_idx).value = header
        ws.merge_cells(start_row=start_row, start_column=col_idx, end_row=start_row + 1, end_column=col_idx)
    ws.column_dimensions["A"].hidden = True

    # 선택목록
    item_level_ranges = {}
    for idx, item in enumerate(items, start=1):
        col_idx = idx
        list_ws.cell(1, col_idx).value = item.get("item_id", "")
        levels = [clean_text(x) for x in item.get("levels", []) if clean_text(x)]
        for row_idx, level in enumerate(levels, start=2):
            list_ws.cell(row_idx, col_idx).value = level
        if levels:
            col_letter = get_column_letter(col_idx)
            item_level_ranges[item.get("item_id", "")] = f"'선택목록'!${col_letter}$2:${col_letter}${len(levels) + 1}"

    for offset, item in enumerate(items):
        col_idx = item_start_col + offset
        assessment_name = get_assessment_name(item.get("assessment_id", ""))
        ws.cell(start_row, col_idx).value = assessment_name
        ws.cell(start_row + 1, col_idx).value = item.get("name", "")
        ws.cell(3, col_idx).value = item.get("item_id", "")
    ws.row_dimensions[3].hidden = True

    # 같은 수행평가명은 가로 병합
    if items:
        group_start = item_start_col
        previous = clean_text(ws.cell(start_row, item_start_col).value)
        for col_idx in range(item_start_col + 1, item_start_col + len(items) + 1):
            current = clean_text(ws.cell(start_row, col_idx).value) if col_idx < item_start_col + len(items) else "__END__"
            if current != previous:
                if previous and col_idx - group_start > 1:
                    ws.merge_cells(start_row=start_row, start_column=group_start, end_row=start_row, end_column=col_idx - 1)
                group_start = col_idx
                previous = current

    border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )
    student_fill = PatternFill("solid", fgColor="E5E7EB")
    assessment_fill = PatternFill("solid", fgColor="DBEAFE")
    item_fill = PatternFill("solid", fgColor="FFEDD5")
    body_fill = PatternFill("solid", fgColor="F9FAFB")
    level_fill = PatternFill("solid", fgColor="FFF7ED")

    for row_idx in [start_row, start_row + 1]:
        for col_idx in range(1, last_col + 1):
            cell = ws.cell(row_idx, col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(bold=True, color="111827")
            if col_idx < item_start_col:
                cell.fill = student_fill
            elif row_idx == start_row:
                cell.fill = assessment_fill
                cell.font = Font(bold=True, color="1E3A8A")
            else:
                cell.fill = item_fill
                cell.font = Font(bold=True, color="92400E")

    for row_offset, (_, student) in enumerate(students.iterrows()):
        row_idx = data_start_row + row_offset
        sid = student.get("student_id", "")
        values = [sid, student.get("학년", ""), student.get("반", ""), student.get("번호", ""), student.get("성명", "")]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = value
            cell.border = border
            cell.fill = body_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for offset, item in enumerate(items):
            col_idx = item_start_col + offset
            cell = ws.cell(row_idx, col_idx)
            cell.value = get_level(sid, item.get("item_id", ""))
            cell.border = border
            cell.fill = level_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = "F6"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 14
    for col_idx in range(item_start_col, last_col + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def export_final_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "최종생기부"
    ws.sheet_properties.tabColor = "FF2563EB"

    headers = ["학년", "반", "번호", "성명", "최종 생기부", "byte", "생성일시", "생성 원문", "API 입력자료"]
    ws.append(headers)
    students = st.session_state.mid_students.copy()
    for _, student in students.iterrows():
        sid = student.get("student_id", "")
        result = st.session_state.mid_results.get(sid, {})
        final_text = result.get("edited", result.get("generated", ""))
        ws.append([
            student.get("학년", ""),
            student.get("반", ""),
            student.get("번호", ""),
            student.get("성명", ""),
            final_text,
            byte_count(final_text),
            result.get("created_at", ""),
            result.get("generated", ""),
            result.get("material", ""),
        ])

    border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )
    header_fill = PatternFill("solid", fgColor="1E3A8A")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.font = Font(size=10, color="111827")
    ws.freeze_panes = "E2"
    ws.sheet_view.showGridLines = False
    widths = {"A": 8, "B": 8, "C": 8, "D": 14, "E": 72, "F": 10, "G": 20, "H": 56, "I": 72}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # 학생별 성취수준 시트: 수행평가명/영역명 2단 헤더
    ws2 = wb.create_sheet("학생별성취수준")
    items = all_used_items()
    base_headers = ["student_id", "학년", "반", "번호", "성명"]
    for col_idx, header in enumerate(base_headers, start=1):
        ws2.cell(1, col_idx).value = header
        ws2.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)
    start_col = 6
    for offset, item in enumerate(items):
        col_idx = start_col + offset
        ws2.cell(1, col_idx).value = get_assessment_name(item.get("assessment_id", ""))
        ws2.cell(2, col_idx).value = item.get("name", "")
    if items:
        group_start = start_col
        previous = clean_text(ws2.cell(1, start_col).value)
        for col_idx in range(start_col + 1, start_col + len(items) + 1):
            current = clean_text(ws2.cell(1, col_idx).value) if col_idx < start_col + len(items) else "__END__"
            if current != previous:
                if previous and col_idx - group_start > 1:
                    ws2.merge_cells(start_row=1, start_column=group_start, end_row=1, end_column=col_idx - 1)
                group_start = col_idx
                previous = current
    for row_offset, (_, student) in enumerate(students.iterrows(), start=3):
        sid = student.get("student_id", "")
        values = [sid, student.get("학년", ""), student.get("반", ""), student.get("번호", ""), student.get("성명", "")]
        for col_idx, value in enumerate(values, start=1):
            ws2.cell(row_offset, col_idx).value = value
        for offset, item in enumerate(items):
            ws2.cell(row_offset, start_col + offset).value = get_level(sid, item.get("item_id", ""))
    ws2.column_dimensions["A"].hidden = True
    ws2.freeze_panes = "F3"
    ws2.sheet_view.showGridLines = False
    for row in ws2.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if cell.row <= 2:
                cell.font = Font(bold=True, color="111827")
                cell.fill = PatternFill("solid", fgColor="DBEAFE" if cell.column >= 6 else "E5E7EB")
    for col_idx in range(1, max(6 + len(items), 6)):
        ws2.column_dimensions[get_column_letter(col_idx)].width = 16

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output



def render_level_header_preview(items):
    # 웹 입력표 위에 수행평가명/관찰 영역명을 app.py 색상 규칙으로 보여준다.
    items = list(items or [])
    if not items:
        return

    base_headers = ["학년", "반", "번호", "성명"]
    assessment_groups = []
    current_name = None
    current_items = []

    for item in items:
        assessment_name = get_assessment_name(item.get("assessment_id", "")) or "수행평가명 미입력"
        if current_name is None:
            current_name = assessment_name
        if assessment_name != current_name:
            assessment_groups.append((current_name, current_items))
            current_name = assessment_name
            current_items = []
        current_items.append(item)
    if current_name is not None:
        assessment_groups.append((current_name, current_items))

    first_row = "".join([
        f'<th class="mid-student-head" rowspan="2">{html.escape(header)}</th>'
        for header in base_headers
    ])
    for assessment_name, group_items in assessment_groups:
        first_row += (
            f'<th class="mid-assessment-head" colspan="{len(group_items)}">'
            f'📁 {html.escape(assessment_name)}</th>'
        )

    second_row = "".join([
        f'<th class="mid-item-head">🧾 {html.escape(clean_text(item.get("name", "이름 없는 관찰 영역")))}</th>'
        for _, group_items in assessment_groups
        for item in group_items
    ])

    html_block = (
        '<div class="mid-color-guide">'
        '<span class="mid-color-chip"><span class="mid-blue-dot"></span>수행평가명</span>'
        '<span class="mid-color-chip"><span class="mid-yellow-dot"></span>관찰 영역명</span>'
        '</div>'
        '<div class="mid-header-preview-wrap"><table class="mid-header-preview"><thead>'
        f'<tr>{first_row}</tr><tr>{second_row}</tr>'
        '</thead></table></div>'
    )
    st.markdown("".join(html_block), unsafe_allow_html=True)

# =========================
# UI 스타일 / 이동
# =========================
st.markdown('<div id="mid-honey-top"></div>', unsafe_allow_html=True)
st.title(MID_APP_TITLE)
st.caption(MID_APP_SUBTITLE)

st.markdown(
    """
    <style>
    div[data-testid="stRadio"] > div[role="radiogroup"] {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 0 !important;
        border-bottom: 1px solid #D0D7DE !important;
        margin-bottom: 1.15rem !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label {
        min-height: 42px !important;
        padding: 0.62rem 0.95rem !important;
        margin-right: 4px !important;
        margin-bottom: -1px !important;
        border: 1px solid #D0D7DE !important;
        border-bottom: none !important;
        border-radius: 10px 10px 0 0 !important;
        background: #F6F8FA !important;
        cursor: pointer !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) {
        background: #FFFFFF !important;
        border-top: 3px solid #D92D20 !important;
        color: #111827 !important;
        font-weight: 800 !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label p {
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] input {
        display: none !important;
    }
    /* app.py와 같은 색상 규칙: 수행평가는 파란색, 관찰 영역은 노란색/주황색 */
    /* 하위 박스: 관찰 영역은 노란색/주황색 계열 */
    div[data-testid="stExpander"] details:has(.mid-item-card-content) {
        background: linear-gradient(180deg, #FFFBF5 0%, #FFF7ED 100%) !important;
        border: 2px solid #FED7AA !important;
        border-radius: 16px !important;
        box-shadow: 0 4px 12px rgba(234, 88, 12, 0.07) !important;
        padding: 0.12rem 0.3rem 0.32rem 0.3rem !important;
        margin: 0.65rem 0 0.95rem 0 !important;
    }
    div[data-testid="stExpander"] details:has(.mid-item-card-content) > summary {
        background: #FFEDD5 !important;
        border: 1px solid #FED7AA !important;
        border-radius: 12px !important;
        margin: 0.16rem 0 0.5rem 0 !important;
        padding: 0.12rem 0.4rem !important;
    }
    div[data-testid="stExpander"] details:has(.mid-item-card-content) > summary p {
        color: #111827 !important;
        font-weight: 800 !important;
    }

    /* 상위 박스: 수행평가는 파란색 계열. 아래 선언이 나중에 와야 중첩 expander에서도 파란색이 유지된다. */
    div[data-testid="stExpander"] details:has(.mid-assessment-card-content) {
        background: linear-gradient(180deg, #EFF6FF 0%, #DBEAFE 100%) !important;
        border: 2px solid #93C5FD !important;
        border-radius: 18px !important;
        box-shadow: 0 5px 15px rgba(37, 99, 235, 0.10) !important;
        padding: 0.18rem 0.38rem 0.42rem 0.38rem !important;
        margin: 0.9rem 0 1.25rem 0 !important;
    }
    div[data-testid="stExpander"] details:has(.mid-assessment-card-content) > summary {
        background: #DBEAFE !important;
        border: 1px solid #BFDBFE !important;
        border-radius: 14px !important;
        margin: 0.2rem 0 0.55rem 0 !important;
        padding: 0.15rem 0.45rem !important;
    }
    div[data-testid="stExpander"] details:has(.mid-assessment-card-content) > summary p {
        color: #0F172A !important;
        font-weight: 900 !important;
    }
    .mid-assessment-card-content, .mid-item-card-content { display: none !important; }

    .mid-header-preview-wrap {
        width: 100%;
        overflow-x: auto;
        border: 1px solid #D1D5DB;
        border-radius: 14px;
        margin: 0.6rem 0 0 0;
        background: #FFFFFF;
        border-bottom: none;
        border-radius: 14px 14px 0 0;
    }
    table.mid-header-preview {
        border-collapse: separate;
        border-spacing: 0;
        min-width: 760px;
        width: 100%;
        table-layout: fixed;
        font-size: 0.92rem;
    }
    table.mid-header-preview th {
        border-right: 1px solid #D1D5DB;
        border-bottom: 1px solid #D1D5DB;
        padding: 9px 10px;
        text-align: center;
        vertical-align: middle;
        line-height: 1.35;
        word-break: keep-all;
    }
    table.mid-header-preview th.mid-student-head {
        background: #E5E7EB;
        color: #111827;
        font-weight: 800;
        width: 72px;
    }
    table.mid-header-preview th.mid-assessment-head {
        background: #DBEAFE;
        color: #1E3A8A;
        font-weight: 900;
    }
    table.mid-header-preview th.mid-item-head {
        background: #FFEDD5;
        color: #92400E;
        font-weight: 900;
    }
    .mid-color-guide {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 0.25rem 0 0.35rem 0;
        font-size: 0.9rem;
    }
    .mid-color-chip {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        border: 1px solid #D1D5DB;
        border-radius: 999px;
        padding: 5px 10px;
        background: #FFFFFF;
        font-weight: 800;
    }
    .mid-blue-dot, .mid-yellow-dot {
        width: 12px;
        height: 12px;
        border-radius: 999px;
        display: inline-block;
    }
    .mid-blue-dot { background: #DBEAFE; border: 1px solid #93C5FD; }
    .mid-yellow-dot { background: #FFEDD5; border: 1px solid #FED7AA; }
.mid-header-preview-wrap + div[data-testid="stDataFrame"],
    .mid-header-preview-wrap + div[data-testid="stDataEditor"] {
        margin-top: 0 !important;
    }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
        margin-top: 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("작업 관리")

    uploaded_project = st.file_uploader(
        "프로젝트 JSON 불러오기",
        type=["json"],
        help="중학교 간편 생기부 프로젝트 파일을 다시 불러옵니다.",
    )
    if uploaded_project and st.button("프로젝트 불러오기"):
        load_mid_project_json(uploaded_project)
        st.success("프로젝트를 불러왔습니다.")
        st.rerun()

    st.download_button(
        "현재 프로젝트 JSON 저장",
        data=mid_project_to_json(),
        file_name="middle_개꿀생기부_project.json",
        mime="application/json",
    )

    st.divider()

    if st.button("샘플 데이터 불러오기", help="중학교 간편 생기부 샘플 프로젝트를 불러옵니다."):
        load_mid_sample_data()
        st.success("샘플 데이터를 불러왔습니다.")
        st.rerun()

    if st.button("전체 초기화", type="secondary"):
        for key in [
            "mid_settings",
            "mid_assessments",
            "mid_items",
            "mid_students",
            "mid_records",
            "mid_results",
            "mid_selected_result_student_id",
            "mid_current_step",
        ]:
            if key in st.session_state:
                del st.session_state[key]
        init_mid_state()
        sanitize_mid_state()
        st.success("초기화했습니다.")
        st.rerun()


STEP_LABELS = [
    "① 기본 설정/수행평가 설계",
    "② 학생별 성취수준 입력",
    "③ 생기부 생성/다운로드",
]
NAV_WIDGET_KEY = "mid_step_nav_radio_v05"
PENDING_STEP_KEY = "mid_pending_step_index_v05"
SCROLL_TO_TOP_KEY = "mid_scroll_to_top_after_step_change_v05"

if "mid_current_step" not in st.session_state:
    st.session_state.mid_current_step = 0

programmatic_step_change = False
if PENDING_STEP_KEY in st.session_state:
    try:
        st.session_state.mid_current_step = int(st.session_state[PENDING_STEP_KEY])
    except Exception:
        st.session_state.mid_current_step = 0
    del st.session_state[PENDING_STEP_KEY]
    st.session_state[SCROLL_TO_TOP_KEY] = True
    programmatic_step_change = True

st.session_state.mid_current_step = max(0, min(int(st.session_state.mid_current_step), len(STEP_LABELS) - 1))
if NAV_WIDGET_KEY not in st.session_state or programmatic_step_change:
    st.session_state[NAV_WIDGET_KEY] = STEP_LABELS[st.session_state.mid_current_step]


def request_step_change(next_index: int):
    st.session_state[PENDING_STEP_KEY] = int(next_index)


def scroll_page_to_top_once():
    if not st.session_state.get(SCROLL_TO_TOP_KEY, False):
        return
    st.session_state[SCROLL_TO_TOP_KEY] = False
    components.html(
        """
        <script>
        function forceScrollTop() {
            try {
                const parentWindow = window.parent;
                const parentDoc = parentWindow.document;
                try { parentWindow.scrollTo(0, 0); } catch (e) {}
                try { parentDoc.documentElement.scrollTop = 0; } catch (e) {}
                try { parentDoc.body.scrollTop = 0; } catch (e) {}
                parentDoc.querySelectorAll('section, main, div').forEach(function(el) {
                    try {
                        if (el.scrollHeight > el.clientHeight + 80) { el.scrollTop = 0; }
                    } catch (e) {}
                });
            } catch (e) {}
        }
        forceScrollTop();
        setTimeout(forceScrollTop, 80);
        setTimeout(forceScrollTop, 250);
        setTimeout(forceScrollTop, 600);
        </script>
        """,
        height=0,
    )


st.markdown("### 작업 단계")
selected_step_label = st.radio(
    "이동할 단계를 선택하세요.",
    STEP_LABELS,
    horizontal=True,
    label_visibility="collapsed",
    key=NAV_WIDGET_KEY,
)
st.session_state.mid_current_step = STEP_LABELS.index(selected_step_label)
current_step = st.session_state.mid_current_step
scroll_page_to_top_once()


def render_next_step_button(current_index: int):
    if current_index >= len(STEP_LABELS) - 1:
        return
    next_index = current_index + 1
    next_label = STEP_LABELS[next_index]
    st.divider()
    st.button(
        f"다음 단계로 넘어가기 → {next_label}",
        type="primary",
        use_container_width=True,
        key=f"mid_next_step_button_{current_index}",
        on_click=request_step_change,
        args=(next_index,),
    )


# =========================
# ① 기본 설정/수행평가 설계
# =========================
if current_step == 0:
    st.subheader("① 기본 설정 / 수행평가 설계")

    settings = st.session_state.mid_settings

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
        settings["grade"] = st.text_input("학년", value=clean_text(settings.get("grade", "2")))
    with col4:
        settings["subject"] = st.text_input("과목명", value=clean_text(settings.get("subject", "과학")))

    col5, col6 = st.columns(2)
    with col5:
        settings["target_bytes_min"] = st.number_input(
            "목표 최소 byte",
            min_value=100,
            max_value=2000,
            value=int(settings.get("target_bytes_min", 250)),
            step=50,
        )
    with col6:
        settings["target_bytes_max"] = st.number_input(
            "목표 최대 byte",
            min_value=100,
            max_value=2000,
            value=int(settings.get("target_bytes_max", 450)),
            step=50,
        )

    st.markdown("#### 생기부 작성 규칙")
    settings["custom_rules"] = st.text_area(
        "공통 작성 규칙",
        value=settings.get("custom_rules", MID_DEFAULT_RULES),
        height=150,
    )

    st.divider()
    st.markdown("### 📚 수행평가와 관찰 영역 설계")
    st.caption("중학교용은 수행평가 안에 관찰 영역을 만들고, 각 영역별 성취수준 코드와 평가 문구만 정해두면 됩니다.")

    with st.expander("➕ 새 수행평가 추가", expanded=True):
        with st.form("mid_add_assessment_form_v05"):
            col_a, col_b = st.columns([2, 1])
            with col_a:
                new_name = st.text_input("수행평가명", placeholder="예: 소화 기관 모형 만들기")
                new_unit = st.text_input("단원", placeholder="예: 동물의 몸과 영양소")
            with col_b:
                new_use = st.checkbox("사용", value=True)
                st.caption("순서는 추가 순서대로 표시됩니다.")
            new_desc = st.text_area("활동 설명", placeholder="예: 소화 기관의 구조와 역할을 모형으로 표현하고 설명하는 활동", height=80)
            submitted = st.form_submit_button("수행평가 추가")
            if submitted:
                if not clean_text(new_name):
                    st.warning("수행평가명을 입력하세요.")
                else:
                    st.session_state.mid_assessments.append({
                        "assessment_id": make_id("mid_assess"),
                        "name": clean_text(new_name),
                        "unit": clean_text(new_unit),
                        "description": clean_text(new_desc),
                        "order": len(st.session_state.mid_assessments) + 1,
                        "use": bool(new_use),
                    })
                    sanitize_mid_state()
                    st.success("수행평가를 추가했습니다.")
                    st.rerun()

    rubric_updates = {}
    assessments = sort_assessments()
    if not assessments:
        st.info("아직 수행평가가 없습니다. 위에서 수행평가를 먼저 추가하세요.")

    for assess_index, assessment in enumerate(assessments, start=1):
        aid = assessment.get("assessment_id", "")
        items = sort_items_for_assessment(aid)
        status_badge = "사용" if assessment.get("use", True) else "미사용"
        title = f"📁 수행평가 {assess_index}. {assessment.get('name', '이름 없는 수행평가')} · 관찰 영역 {len(items)}개 · {status_badge}"
        with st.expander(title, expanded=True):
            st.markdown('<div class="mid-assessment-card-content"></div>', unsafe_allow_html=True)
            st.markdown(f"### 📁 수행평가 {assess_index}. {assessment.get('name', '')}")
            if assessment.get("unit"):
                st.caption(f"단원: {assessment.get('unit', '')}")
            if assessment.get("description"):
                st.caption(f"활동 설명: {assessment.get('description', '')}")

            with st.expander("⚙️ 수행평가 기본 정보 수정 / 삭제", expanded=False):
                col_x, col_y, col_z = st.columns([2, 2, 1])
                with col_x:
                    assessment["name"] = st.text_input("수행평가명 수정", value=assessment.get("name", ""), key=f"mid_assess_name_{aid}")
                    assessment["unit"] = st.text_input("단원 수정", value=assessment.get("unit", ""), key=f"mid_assess_unit_{aid}")
                with col_y:
                    assessment["description"] = st.text_area("활동 설명 수정", value=assessment.get("description", ""), key=f"mid_assess_desc_{aid}", height=100)
                with col_z:
                    assessment["use"] = st.checkbox("사용", value=assessment.get("use", True), key=f"mid_assess_use_{aid}")
                    if st.button("수행평가 삭제", key=f"mid_delete_assessment_{aid}"):
                        item_ids = [item.get("item_id", "") for item in get_items_for_assessment(aid)]
                        st.session_state.mid_assessments = [a for a in st.session_state.mid_assessments if a.get("assessment_id", "") != aid]
                        st.session_state.mid_items = [item for item in st.session_state.mid_items if item.get("assessment_id", "") != aid]
                        st.session_state.mid_records = {k: v for k, v in as_dict(st.session_state.mid_records).items() if k.split("::")[-1] not in item_ids}
                        st.success("수행평가를 삭제했습니다.")
                        st.rerun()

            st.markdown("#### 🧾 관찰 영역")
            if not items:
                st.info("아직 관찰 영역이 없습니다. 아래에서 관찰 영역을 추가하세요.")

            for item_index, item in enumerate(items, start=1):
                item_id = item.get("item_id", "")
                with st.expander(f"🧾 관찰 영역 {item_index}. {item.get('name', '이름 없는 관찰 영역')}", expanded=True):
                    st.markdown('<div class="mid-item-card-content"></div>', unsafe_allow_html=True)
                    col_i1, col_i2 = st.columns([4, 1])
                    with col_i1:
                        item["name"] = st.text_input("관찰 영역명", value=item.get("name", ""), key=f"mid_item_name_{item_id}")
                    with col_i2:
                        st.caption(f"현재 {item_index}번째")
                        if st.button("관찰 영역 삭제", key=f"mid_delete_item_{item_id}"):
                            st.session_state.mid_items = [x for x in st.session_state.mid_items if x.get("item_id", "") != item_id]
                            st.session_state.mid_records = {k: v for k, v in as_dict(st.session_state.mid_records).items() if not k.endswith(f"::{item_id}")}
                            st.success("관찰 영역을 삭제했습니다.")
                            st.rerun()

                    current_levels = item.get("levels", []) if isinstance(item.get("levels", []), list) else []
                    current_rubrics = item.get("rubrics", {}) if isinstance(item.get("rubrics", {}), dict) else {}
                    default_count = max(1, min(10, len(current_levels) if current_levels else 3))
                    level_count = st.selectbox(
                        "성취수준 개수",
                        options=list(range(1, 11)),
                        index=default_count - 1,
                        key=f"mid_level_count_{item_id}",
                    )
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
                                key=f"mid_code_{item_id}_{idx}",
                                label_visibility="collapsed",
                                placeholder="A",
                            )
                        with col_text:
                            comment = st.text_area(
                                f"{idx + 1}번 평가 문구",
                                value=existing_text,
                                key=f"mid_rubric_{item_id}_{idx}",
                                height=70,
                                label_visibility="collapsed",
                                placeholder="예: 핵심 개념을 정확히 연결하여 활동 결과를 구체적으로 설명함",
                            )
                        code = clean_text(code)
                        if code:
                            levels.append(code)
                            rubrics[code] = clean_text(comment)
                    rubric_updates[item_id] = {"levels": levels, "rubrics": rubrics}

            st.divider()
            with st.expander("➕ 이 수행평가에 관찰 영역 추가", expanded=(len(items) == 0)):
                new_item_name = st.text_input("관찰 영역명", placeholder="예: 소화 기관의 순서와 역할", key=f"mid_new_item_name_{aid}")
                if st.button("이 수행평가에 관찰 영역 추가", key=f"mid_add_item_button_{aid}"):
                    if not clean_text(new_item_name):
                        st.warning("관찰 영역명을 입력하세요.")
                    else:
                        st.session_state.mid_items.append({
                            "item_id": make_id("mid_item"),
                            "assessment_id": aid,
                            "name": clean_text(new_item_name),
                            "levels": ["A", "B", "C"],
                            "rubrics": {
                                "A": "핵심 개념을 정확히 연결하여 활동 결과를 구체적으로 설명함",
                                "B": "주요 개념을 바탕으로 활동 결과를 대체로 적절하게 설명함",
                                "C": "기초 개념을 중심으로 활동에 참여하고 기본 내용을 정리함",
                            },
                            "order": len(items) + 1,
                        })
                        sanitize_mid_state()
                        st.success("관찰 영역을 추가했습니다.")
                        st.rerun()

    if rubric_updates:
        st.divider()
        if st.button("전체 성취수준/평가 문구 한꺼번에 저장", type="primary", use_container_width=True):
            saved_count = 0
            for item in st.session_state.mid_items:
                item_id = item.get("item_id", "")
                if item_id in rubric_updates:
                    item["levels"] = rubric_updates[item_id]["levels"]
                    item["rubrics"] = rubric_updates[item_id]["rubrics"]
                    saved_count += 1
            sanitize_mid_state()
            st.success(f"성취수준/평가 문구를 {saved_count}개 관찰 영역에 저장했습니다.")
            st.rerun()

    render_next_step_button(0)


# =========================
# ② 학생별 성취수준 입력
# =========================
if current_step == 1:
    st.subheader("② 학생별 성취수준 입력")

    items = all_used_items()
    if not items:
        st.warning("먼저 ①에서 수행평가와 관찰 영역을 추가하세요.")
    else:
        st.markdown(
            """
            학생은 행으로, 관찰 영역은 열로 입력합니다.  
            영역 열은 **수행평가명 / 관찰 영역명** 형태로 표시됩니다. 이름은 관리용이며 AI 프롬프트에는 들어가지 않습니다.
            """
        )

        matrix_df, items = build_record_matrix_df()
        visible_df = matrix_df.copy()

        column_config = {
            "student_id": None,
            "학년": st.column_config.TextColumn("학년", width="small"),
            "반": st.column_config.TextColumn("반", width="small"),
            "번호": st.column_config.TextColumn("번호", width="small"),
            "성명": st.column_config.TextColumn("성명", width="medium"),
        }
        for item in items:
            item_id = item.get("item_id", "")
            assessment_name = get_assessment_name(item.get("assessment_id", "")) or "수행평가명 미입력"
            item_name = clean_text(item.get("name", "")) or "관찰 영역명 미입력"
            # Streamlit 입력표 안에서 바로 구분되도록 열 제목을 2줄로 표시한다.
            # 병합 헤더는 아니지만, 각 열마다 위 줄은 수행평가명, 아래 줄은 관찰 영역명으로 반복된다.
            label = f"📁 {assessment_name}\n🧾 {item_name}"
            levels = [clean_text(x) for x in item.get("levels", []) if clean_text(x)]
            column_config[item_id] = st.column_config.SelectboxColumn(
                label,
                options=[""] + levels,
                required=False,
                width="medium",
                help=f"{assessment_name} / {item_name}",
            )

        st.markdown("#### 학생별 성취수준 입력표")
        st.caption("각 영역 열 제목을 2줄로 표시했습니다. 위 줄은 수행평가명, 아래 줄은 관찰 영역명입니다.")

        edited_df = st.data_editor(
            visible_df,
            num_rows="dynamic",
            use_container_width=True,
            height=560,
            key="mid_record_matrix_editor_v07",
            column_config=column_config,
        )

        col_save, col_download = st.columns([1.4, 2.0])
        with col_save:
            if st.button("학생별 성취수준 저장", type="primary", use_container_width=True):
                saved_count = save_record_matrix_df(edited_df, items)
                st.success(f"학생별 성취수준을 저장했습니다. 저장된 입력값: {saved_count}개")
                st.rerun()
        with col_download:
            st.download_button(
                "학생별 성취수준 입력표 엑셀 다운로드",
                data=make_record_input_excel(),
                file_name=f"middle_student_levels_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with st.expander("표 헤더 구조 설명", expanded=False):
            st.markdown(
                "웹 입력표는 병합 헤더 대신 각 성취수준 열 제목에 수행평가명과 관찰 영역명을 2줄로 반복 표시합니다. "
                "그래서 별도 헤더 미리보기 없이 실제 입력표 안에서 바로 구분할 수 있습니다. "
                "엑셀 다운로드 파일에서는 윗줄 수행평가명, 아랫줄 관찰 영역명의 2단 헤더를 적용했습니다."
            )

    render_next_step_button(1)


# =========================
# ③ 생기부 생성/다운로드
# =========================
if current_step == 2:
    st.subheader("③ 생기부 생성 / 수정 / 다운로드")

    if st.session_state.mid_students.empty:
        st.warning("먼저 ②에서 학생별 성취수준을 입력하고 저장하세요.")
    elif not all_used_items():
        st.warning("먼저 ①에서 수행평가와 관찰 영역을 추가하세요.")
    else:
        st.markdown("#### AI 선택 및 API 설정")
        col_provider, col_model, col_key, col_variation = st.columns([1.2, 1.7, 2.4, 1.2], gap="small")
        with col_provider:
            ai_provider = st.selectbox("사용할 AI", AI_PROVIDER_OPTIONS, index=0)
        with col_model:
            model = st.text_input(
                "모델명",
                value=get_default_ai_model(ai_provider),
                key=f"mid_ai_model_name_{ai_provider}_v06",
            )
        with col_key:
            api_key = st.text_input(
                f"{ai_provider} API Key",
                value=get_default_ai_key(ai_provider),
                type="password",
                key=f"mid_api_key_{ai_provider}_v06",
                help=f"Streamlit Secrets에는 {AI_SECRET_KEY_NAMES.get(ai_provider, 'OPENAI_API_KEY')} 이름으로 저장해둘 수 있습니다.",
            )
        with col_variation:
            variation_level = st.selectbox("변주 강도", ["낮음", "보통", "높음"], index=1)

        students = st.session_state.mid_students.copy()
        student_labels = {
            f"{row.get('학년', '')}학년 {row.get('반', '')}반 {row.get('번호', '')}번 {row.get('성명', '')}": row
            for _, row in students.iterrows()
        }
        selected_label = st.selectbox("생성할 학생", list(student_labels.keys()))
        selected_student = student_labels[selected_label]

        col_a, col_b, col_spacer = st.columns([1.65, 1.9, 5.45], gap="small")
        with col_a:
            if st.button("선택 학생 생기부 생성", type="secondary", use_container_width=True):
                sid = selected_student.get("student_id", "")
                variant_no = list(students["student_id"].astype(str)).index(str(sid)) + 1 if sid in students["student_id"].astype(str).tolist() else 1
                material, generated = generate_for_student(selected_student, ai_provider, api_key, model, variation_level, variant_no)
                st.session_state.mid_results[sid] = {
                    "material": material,
                    "generated": generated,
                    "edited": generated,
                    "bytes": byte_count(generated),
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                st.success("생성했습니다.")
                st.rerun()
        with col_b:
            if st.button("전체 학생 생기부 생성 시작", type="primary", use_container_width=True):
                progress = st.progress(0)
                total = len(students)
                for idx, (_, student) in enumerate(students.iterrows(), start=1):
                    sid = student.get("student_id", "")
                    material, generated = generate_for_student(student, ai_provider, api_key, model, variation_level, idx)
                    st.session_state.mid_results[sid] = {
                        "material": material,
                        "generated": generated,
                        "edited": generated,
                        "bytes": byte_count(generated),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    progress.progress(idx / total)
                st.success("전체 학생 생기부 생성을 완료했습니다.")
                st.rerun()

        st.divider()
        st.markdown("#### 전체 결과표 / 학생 선택")
        result_rows = []
        for _, student in students.iterrows():
            sid = student.get("student_id", "")
            result = st.session_state.mid_results.get(sid, {})
            text = result.get("edited", result.get("generated", ""))
            result_rows.append({
                "student_id": sid,
                "학년": student.get("학년", ""),
                "반": student.get("반", ""),
                "번호": student.get("번호", ""),
                "성명": student.get("성명", ""),
                "생성/수정 문구": text,
                "byte": byte_count(text),
                "생성일시": result.get("created_at", ""),
            })
        result_df = pd.DataFrame(result_rows)
        display_result_df = result_df.drop(columns=["student_id"], errors="ignore")

        if clean_text(st.session_state.get("mid_selected_result_student_id", "")) not in set(students["student_id"].astype(str)):
            st.session_state.mid_selected_result_student_id = selected_student.get("student_id", "")
        current_selected_sid = clean_text(st.session_state.mid_selected_result_student_id)

        selection_event = st.dataframe(
            display_result_df,
            use_container_width=True,
            height=340,
            hide_index=True,
            key="mid_generation_result_selector_v05",
            on_select="rerun",
            selection_mode="single-cell",
            column_config={
                "학년": st.column_config.TextColumn("학년", width="small"),
                "반": st.column_config.TextColumn("반", width="small"),
                "번호": st.column_config.TextColumn("번호", width="small"),
                "성명": st.column_config.TextColumn("성명", width="medium"),
                "생성/수정 문구": st.column_config.TextColumn("생성/수정 문구", width="large"),
                "byte": st.column_config.NumberColumn("byte", width="small"),
                "생성일시": st.column_config.TextColumn("생성일시", width="medium"),
            },
        )
        selected_row_index = None
        try:
            selected_cells = list(selection_event.selection.cells)
            if selected_cells:
                selected_row_index = int(selected_cells[-1][0])
        except Exception:
            selected_row_index = None
        if selected_row_index is not None and 0 <= selected_row_index < len(result_df):
            current_selected_sid = clean_text(result_df.iloc[selected_row_index].get("student_id", current_selected_sid))
            st.session_state.mid_selected_result_student_id = current_selected_sid

        selected_match = students[students["student_id"] == current_selected_sid]
        if selected_match.empty:
            selected_detail_student = selected_student
            current_selected_sid = selected_student.get("student_id", "")
        else:
            selected_detail_student = selected_match.iloc[0]
        selected_detail_label = f"{selected_detail_student.get('학년', '')}학년 {selected_detail_student.get('반', '')}반 {selected_detail_student.get('번호', '')}번 {selected_detail_student.get('성명', '')}"
        result = st.session_state.mid_results.get(current_selected_sid, {})
        initial_text = result.get("edited", result.get("generated", ""))

        st.markdown(f"#### 표에서 선택한 학생 수정: {selected_detail_label}")
        if not result:
            st.info("아직 이 학생의 생기부 문구가 생성되지 않았습니다. 직접 입력하거나 위에서 생성을 먼저 실행하세요.")
        editor_key = f"mid_generation_detail_text_v05_{current_selected_sid}"
        if editor_key not in st.session_state:
            st.session_state[editor_key] = initial_text
        edited = st.text_area("교사 수정 문구", key=editor_key, height=200)
        st.metric("현재 byte", byte_count(edited))
        if st.button("수정 문구 저장", type="primary"):
            result = st.session_state.mid_results.get(current_selected_sid, {})
            result["edited"] = edited
            result["bytes"] = byte_count(edited)
            if not result.get("generated"):
                result["generated"] = edited
            if not result.get("material"):
                result["material"] = build_student_material(selected_detail_student)
            if not result.get("created_at"):
                result["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.mid_results[current_selected_sid] = result
            st.success("수정 문구를 저장했습니다.")
            st.rerun()

        st.divider()
        st.markdown("#### 최종 결과 다운로드")
        st.caption("마지막 단계입니다. 수정 내용을 확인한 뒤 결과 엑셀을 내려받으세요.")
        st.markdown(
            """
            <style>
            section.main div[data-testid="stDownloadButton"] > button {
                min-height: 58px !important;
                font-size: 1.14rem !important;
                font-weight: 900 !important;
                border-radius: 14px !important;
                box-shadow: 0 4px 12px rgba(217, 45, 32, 0.18) !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            "📥 결과 엑셀 다운로드",
            data=export_final_excel(),
            file_name=f"middle_개꿀생기부_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
