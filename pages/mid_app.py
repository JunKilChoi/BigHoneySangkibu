import hashlib
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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from streamlit_sortables import sort_items
except Exception:
    sort_items = None


# =========================
# 중학교 간편 생성기 기본 설정
# =========================
st.set_page_config(
    page_title="중학교 간편 생기부",
    page_icon="🍯",
    layout="wide",
)

MID_APP_TITLE = "🍯 중학교 간편 생기부 v02"
MID_APP_SUBTITLE = "수행평가 기반 중학교 생기부 문장 변주 도우미 · patched-20260623-mid-v02"
MID_APP_VERSION = "patched-20260623-mid-v02"

MID_DEFAULT_RULES = """- 중학교 학교생활기록부 교과 세부능력 및 특기사항 문체로 작성한다.
- 학생 이름, 학년, 반, 번호, 학교명 등 개인정보를 쓰지 않는다.
- 첫 문장을 '학생은', '이 학생은', '해당 학생은'으로 시작하지 않는다.
- 활동 결과에 없는 내용을 추측하거나 과장하지 않는다.
- '깊은 이해', '창의융합', '혁신적', '흥미와 전문성 심화', '본인은', '의지를 밝힘' 같은 표현을 피한다.
- 한 문장 또는 짧은 한 문단으로 작성한다.
- 명사형 종결을 사용한다. 예: 수행함, 설명함, 정리함, 제시함, 해석함, 이해한 것으로 보임."""

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

VARIATION_OPTIONS = {
    "낮음": "문장 구조는 크게 바꾸지 말고 어휘와 순서만 조금씩 바꾼다.",
    "보통": "의미는 유지하되 문장 구조와 표현을 적당히 다르게 만든다.",
    "높음": "활동 결과의 의미는 유지하면서 문장 흐름과 표현을 다양하게 만든다.",
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


def get_default_ai_model(provider: str) -> str:
    provider = provider if provider in AI_DEFAULT_MODELS else "ChatGPT"
    return AI_DEFAULT_MODELS[provider]


def get_default_ai_key(provider: str) -> str:
    secret_name = AI_SECRET_KEY_NAMES.get(provider, "OPENAI_API_KEY")
    try:
        return st.secrets.get(secret_name, "")
    except Exception:
        return ""


def default_level_code(index: int) -> str:
    default_codes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    if 0 <= index < len(default_codes):
        return default_codes[index]
    return str(index + 1)


def normalize_sentence(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip(" \n\t-•·")
    text = re.sub(r"^[0-9]+[\.)]\s*", "", text)
    text = text.strip('"\'“”‘’')
    return text


def sort_areas(areas):
    return sorted(
        areas,
        key=lambda x: (
            int(x.get("order", 999) or 999),
            clean_text(x.get("name", "")),
            clean_text(x.get("area_id", "")),
        ),
    )


def area_label(area, idx=None):
    prefix = f"{idx}. " if idx is not None else ""
    name = clean_text(area.get("name", "이름 없는 수행평가")) or "이름 없는 수행평가"
    unit = clean_text(area.get("unit", ""))
    return f"{prefix}{name}" + (f" · {unit}" if unit else "")


def get_area_by_id(area_id):
    for area in st.session_state.get("mid_areas", []):
        if area.get("area_id", "") == area_id:
            return area
    return None


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

    if "mid_areas" not in st.session_state:
        st.session_state.mid_areas = []

    if "mid_results" not in st.session_state:
        st.session_state.mid_results = []

    if "mid_student_rows" not in st.session_state:
        st.session_state.mid_student_rows = pd.DataFrame(
            columns=["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]
        )

    if "mid_selected_result_id" not in st.session_state:
        st.session_state.mid_selected_result_id = ""


def sanitize_mid_state():
    settings = st.session_state.get("mid_settings", {})
    if "custom_rules" not in settings:
        settings["custom_rules"] = MID_DEFAULT_RULES
    if "target_bytes_min" not in settings:
        settings["target_bytes_min"] = 250
    if "target_bytes_max" not in settings:
        settings["target_bytes_max"] = 450
    st.session_state.mid_settings = settings

    clean_areas = []
    for idx, area in enumerate(st.session_state.get("mid_areas", []), start=1):
        if not isinstance(area, dict):
            continue
        if not clean_text(area.get("area_id", "")):
            area["area_id"] = make_id("mid_area")
        area["name"] = clean_text(area.get("name", "")) or "이름 없는 수행평가"
        area["unit"] = clean_text(area.get("unit", ""))
        area["description"] = clean_text(area.get("description", ""))
        area["order"] = int(area.get("order", idx) or idx)
        area["use"] = bool(area.get("use", True))
        if not isinstance(area.get("levels", []), list):
            area["levels"] = []
        if not isinstance(area.get("rubrics", {}), dict):
            area["rubrics"] = {}
        if not area["levels"]:
            area["levels"] = ["A", "B", "C"]
        area["levels"] = [clean_text(x) for x in area["levels"] if clean_text(x)]
        for level in area["levels"]:
            area["rubrics"].setdefault(level, "")
        clean_areas.append(area)

    clean_areas = sort_areas(clean_areas)
    for idx, area in enumerate(clean_areas, start=1):
        area["order"] = idx
    st.session_state.mid_areas = clean_areas

    if not isinstance(st.session_state.get("mid_results"), list):
        st.session_state.mid_results = []

    clean_results = []
    for result in st.session_state.get("mid_results", []):
        if not isinstance(result, dict):
            continue
        if not clean_text(result.get("result_id", "")):
            result["result_id"] = make_id("mid_result")
        result["final_text"] = clean_text(result.get("final_text", result.get("generated_text", "")))
        result["generated_text"] = clean_text(result.get("generated_text", result.get("final_text", "")))
        result["byte"] = byte_count(result.get("final_text", ""))
        clean_results.append(result)
    st.session_state.mid_results = clean_results

    if not isinstance(st.session_state.get("mid_student_rows"), pd.DataFrame):
        st.session_state.mid_student_rows = pd.DataFrame(st.session_state.get("mid_student_rows", []))
    for col in ["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]:
        if col not in st.session_state.mid_student_rows.columns:
            st.session_state.mid_student_rows[col] = "" if col != "생성 수" else 1


init_mid_state()
sanitize_mid_state()


# =========================
# JSON 저장/불러오기 및 샘플
# =========================
def mid_project_to_json() -> str:
    data = {
        "settings": st.session_state.mid_settings,
        "areas": st.session_state.mid_areas,
        "student_rows": st.session_state.mid_student_rows.to_dict(orient="records"),
        "results": st.session_state.mid_results,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "개꿀 생기부 - 중학교 간편 생성기",
        "version": MID_APP_VERSION,
    }
    return json.dumps(json_safe(data), ensure_ascii=False, indent=2, default=str)


def apply_mid_project_data(data):
    st.session_state.mid_settings = data.get("settings", st.session_state.mid_settings)
    st.session_state.mid_areas = data.get("areas", [])
    st.session_state.mid_student_rows = pd.DataFrame(data.get("student_rows", []))
    st.session_state.mid_results = data.get("results", [])
    st.session_state.mid_selected_result_id = ""
    sanitize_mid_state()


def load_mid_project_json(uploaded_file):
    data = json.load(uploaded_file)
    apply_mid_project_data(data)


def build_mid_sample_project():
    areas = [
        {
            "area_id": "mid_area_sample_digest",
            "name": "소화 기관 모형 만들기",
            "unit": "동물의 몸과 영양소",
            "description": "소화 기관의 구조와 각 기관의 역할을 모형으로 표현하고, 음식물이 이동하는 순서를 설명하는 활동",
            "order": 1,
            "use": True,
            "levels": ["상", "중", "하"],
            "rubrics": {
                "상": "소화 기관의 순서와 역할을 정확히 연결하여 모형으로 표현함",
                "중": "주요 소화 기관의 위치와 역할을 대체로 설명함",
                "하": "소화 기관의 기본 구조를 중심으로 모형 제작에 참여함",
            },
        },
        {
            "area_id": "mid_area_sample_photo",
            "name": "광합성 조건 탐구하기",
            "unit": "식물과 에너지",
            "description": "빛, 물, 이산화 탄소 등 광합성에 영향을 주는 조건을 정하고 변인을 통제하여 실험을 설계하는 활동",
            "order": 2,
            "use": True,
            "levels": ["A", "B", "C", "D"],
            "rubrics": {
                "A": "실험군과 대조군을 구분하고 변인 통제 조건을 구체적으로 제시함",
                "B": "광합성에 영향을 주는 조건을 정하고 비교 실험의 기본 구조를 설명함",
                "C": "광합성과 관련된 주요 조건을 바탕으로 실험 설계에 참여함",
                "D": "광합성 조건 탐구의 기본 절차를 중심으로 활동을 수행함",
            },
        },
    ]
    student_rows = [
        {"반": "1", "번호": "1", "성명": "김민준", "성취수준": "상", "추가 코멘트": "기관별 역할을 화살표로 연결함", "생성 수": 1},
        {"반": "1", "번호": "2", "성명": "박서연", "성취수준": "중", "추가 코멘트": "위와 작은창자의 역할을 구분함", "생성 수": 1},
        {"반": "1", "번호": "3", "성명": "최도윤", "성취수준": "상", "추가 코멘트": "음식물 이동 순서를 근거와 함께 설명함", "생성 수": 1},
        {"반": "1", "번호": "4", "성명": "이하은", "성취수준": "하", "추가 코멘트": "모형 제작 과정에 참여함", "생성 수": 1},
    ]
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
        "areas": areas,
        "student_rows": student_rows,
        "results": [],
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "개꿀 생기부 - 중학교 간편 생성기",
        "version": "sample-mid-project-v02",
    }


# =========================
# 수행평가 순서 정렬
# =========================
def normalize_area_orders():
    areas = sort_areas(st.session_state.mid_areas)
    for idx, area in enumerate(areas, start=1):
        area["order"] = idx
    st.session_state.mid_areas = areas


def apply_area_drag_order(sorted_labels, label_to_area_id):
    id_to_area = {area.get("area_id", ""): area for area in st.session_state.mid_areas}
    for idx, label in enumerate(sorted_labels, start=1):
        area_id = label_to_area_id.get(label)
        if area_id in id_to_area:
            id_to_area[area_id]["order"] = idx


def sortable_style():
    return """
    .sortable-component {
        width: 100% !important;
        border: 1px solid #D1D5DB !important;
        border-radius: 12px !important;
        padding: 10px !important;
        background-color: #F9FAFB !important;
        box-sizing: border-box !important;
        overflow: visible !important;
    }
    .sortable-container {
        width: 100% !important;
        background-color: #F9FAFB !important;
        border-radius: 10px !important;
        padding: 6px !important;
        box-sizing: border-box !important;
        overflow: visible !important;
    }
    .sortable-container-header {
        width: 100% !important;
        background-color: #F3F4F6 !important;
        color: #111827 !important;
        font-weight: 700 !important;
        padding: 8px 12px !important;
        border-radius: 8px !important;
        border: 1px solid #E5E7EB !important;
        box-sizing: border-box !important;
    }
    .sortable-container-body {
        width: 100% !important;
        display: flex !important;
        flex-direction: column !important;
        flex-wrap: nowrap !important;
        align-items: stretch !important;
        gap: 8px !important;
        background-color: #F9FAFB !important;
        padding-top: 8px !important;
        box-sizing: border-box !important;
        overflow: visible !important;
    }
    .sortable-item,
    .sortable-item:hover,
    .sortable-container-body > div,
    .sortable-container-body > div:hover {
        display: block !important;
        width: 100% !important;
        min-height: 44px !important;
        background-color: #FFFFFF !important;
        color: #111827 !important;
        font-weight: 700 !important;
        border: 1px solid #D1D5DB !important;
        border-radius: 10px !important;
        padding: 11px 14px !important;
        margin: 0 !important;
        box-shadow: 0 1px 2px rgba(17, 24, 39, 0.06) !important;
        box-sizing: border-box !important;
        white-space: normal !important;
        line-height: 1.35 !important;
        overflow-wrap: anywhere !important;
    }
    .sortable-item:hover,
    .sortable-container-body > div:hover {
        background-color: #F3F4F6 !important;
        border-color: #9CA3AF !important;
    }
    .sortable-item::before {
        content: "☰ " !important;
        color: #6B7280 !important;
        font-weight: 700 !important;
    }
    """


def sortable_key(base_key, labels):
    raw = "||".join([clean_text(label) for label in labels])
    digest = hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]
    return f"{base_key}_{len(labels)}_{digest}"


def sort_labels_with_gray_box(labels, key, header="정렬"):
    if sort_items is None or len(labels) < 2:
        return labels
    component_key = sortable_key(key, labels)
    try:
        return sort_items(labels, header=header, custom_style=sortable_style(), key=component_key)
    except TypeError:
        try:
            return sort_items(labels, key=f"{component_key}_basic")
        except Exception as e:
            st.error(f"드래그 정렬 컴포넌트 오류: {e}")
            return labels
    except Exception as e:
        st.error(f"드래그 정렬 컴포넌트 오류: {e}")
        return labels


# =========================
# AI 호출 함수
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
            "max_tokens": 2500,
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
        text = "".join(
            [block.get("text", "") for block in blocks if isinstance(block, dict) and block.get("type") == "text"]
        ).strip()
        return text or None
    except Exception as e:
        st.error(f"Claude API 생성 중 오류가 발생했습니다: {e}")
        return None


def generate_with_ai(prompt, provider, api_key, model):
    provider = provider if provider in AI_PROVIDER_OPTIONS else "ChatGPT"
    model = clean_text(model) or get_default_ai_model(provider)
    if provider == "Gemini":
        return generate_with_gemini(prompt, api_key, model)
    if provider == "Claude":
        return generate_with_claude(prompt, api_key, model)
    return generate_with_openai(prompt, api_key, model)


def parse_ai_sentences(text: str, expected_count: int) -> list:
    raw = clean_text(text)
    if not raw:
        return []

    fenced = re.sub(r"^```(?:json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    try:
        loaded = json.loads(fenced)
        if isinstance(loaded, list):
            return [normalize_sentence(x) for x in loaded if clean_text(x)][:expected_count]
        if isinstance(loaded, dict):
            for key in ["sentences", "results", "문장"]:
                if isinstance(loaded.get(key), list):
                    return [normalize_sentence(x) for x in loaded[key] if clean_text(x)][:expected_count]
    except Exception:
        pass

    lines = []
    for line in raw.splitlines():
        line = normalize_sentence(line)
        if not line:
            continue
        if line.lower() in ["json", "sentences"]:
            continue
        lines.append(line)

    if len(lines) < expected_count:
        split_candidates = re.split(r"(?<=[함임음봄됨])\s*[;/]\s*", raw)
        for candidate in split_candidates:
            candidate = normalize_sentence(candidate)
            if candidate and candidate not in lines:
                lines.append(candidate)

    unique = []
    seen = set()
    for line in lines:
        if line not in seen:
            unique.append(line)
            seen.add(line)
    return unique[:expected_count]


# =========================
# 프롬프트와 폴백 생성
# =========================
def build_mid_generation_prompt(area, level, rubric, count, variation_strength, extra_comment=""):
    settings = st.session_state.mid_settings
    levels = [clean_text(x) for x in area.get("levels", []) if clean_text(x)]
    level_text = f"전체 성취수준 {', '.join(levels)} 중 {level}" if levels else f"성취수준 {level}"
    extra_comment = clean_text(extra_comment)

    prompt = f"""
너는 중학교 학교생활기록부 교과 세부능력 및 특기사항 문장을 작성하는 교사 보조 도구이다.

[작성 목적]
같은 활동과 같은 성취수준에 해당하는 학생들에게 사용할 수 있는 문장을 만든다.
의미와 평가 근거는 유지하되, 문장 표현은 서로 조금씩 다르게 변주한다.

[공통 규칙]
{settings.get('custom_rules', MID_DEFAULT_RULES)}

[활동 및 관찰 영역]
- 과목: {settings.get('subject', '과학')}
- 수행평가/관찰 영역: {area.get('name', '')}
- 단원/영역: {area.get('unit', '')}
- 활동 설명: {area.get('description', '')}
- 성취수준: {level_text}
- 교사의 평가: {rubric}
""".strip()

    if extra_comment:
        prompt += f"\n- 추가 코멘트: {extra_comment}"

    prompt += f"""

[생성 조건]
- 목표 분량: {settings.get('target_bytes_min', 250)}~{settings.get('target_bytes_max', 450)} byte
- 생성 개수: {int(count)}개
- 변주 강도: {variation_strength} / {VARIATION_OPTIONS.get(variation_strength, VARIATION_OPTIONS['보통'])}
- 모든 문장은 같은 성취수준의 평가 근거를 벗어나지 않는다.
- 학생 이름, 학년, 반, 번호는 절대 쓰지 않는다.
- 제목, 번호, 따옴표, 설명 문구 없이 문장만 출력한다.
- 출력은 JSON 배열 형식으로만 한다. 예: ["문장1", "문장2"]
""".strip()
    return prompt


def fallback_variations(area, level, rubric, count, extra_comment=""):
    name = clean_text(area.get("name", "활동")) or "활동"
    unit = clean_text(area.get("unit", ""))
    desc = clean_text(area.get("description", ""))
    rubric = clean_text(rubric) or f"{name}의 기본 내용을 바탕으로 활동을 수행함"
    extra = clean_text(extra_comment)

    base_clauses = [
        f"{name}에서 {rubric}",
        f"{unit} 단원의 {name} 활동에서 {rubric}" if unit else f"{name} 활동에서 {rubric}",
        f"{re.sub(r'[.。]$', '', desc)} 과정에서 {rubric}" if desc else f"{name} 수행 과정에서 {rubric}",
        f"{name} 활동을 통해 {rubric}",
        f"{name}의 수행 과정에서 {rubric}",
        f"{name}과 관련하여 {rubric}",
        f"{name} 활동에서 핵심 내용을 바탕으로 {rubric}",
        f"{name}을 수행하며 {rubric}",
    ]

    results = []
    for idx in range(max(1, int(count))):
        text = base_clauses[idx % len(base_clauses)]
        if extra:
            if idx % 3 == 0:
                text += f". {extra}"
            elif idx % 3 == 1:
                text += f", 특히 {extra}"
            else:
                text += f". 활동 과정에서 {extra}"
        text = normalize_sentence(text)
        if not text.endswith(("함", "임", "음", "봄", "됨")):
            text = text.rstrip(".") + "함"
        results.append(text)

    return results[: int(count)]


def generate_variations(area, level, count, provider, api_key, model, variation_strength, extra_comment=""):
    count = max(1, min(50, int(count)))
    rubric = clean_text(area.get("rubrics", {}).get(level, ""))
    prompt = build_mid_generation_prompt(area, level, rubric, count, variation_strength, extra_comment=extra_comment)

    generated_text = None
    if clean_text(api_key):
        generated_text = generate_with_ai(prompt, provider, api_key, model)

    sentences = parse_ai_sentences(generated_text, count) if generated_text else []
    if len(sentences) < count:
        fallback = fallback_variations(area, level, rubric, count - len(sentences), extra_comment=extra_comment)
        sentences.extend(fallback)

    return sentences[:count], prompt


# =========================
# 엑셀 다운로드
# =========================
def export_mid_excel():
    results = st.session_state.mid_results
    result_df = pd.DataFrame(results)
    if result_df.empty:
        result_df = pd.DataFrame(
            columns=[
                "mode", "area_name", "level", "variant_no", "class", "number", "name", "final_text", "byte", "created_at", "generated_text", "extra_comment"
            ]
        )

    column_rename = {
        "mode": "생성방식",
        "area_name": "수행평가/관찰영역",
        "level": "성취수준",
        "variant_no": "문장번호",
        "class": "반",
        "number": "번호",
        "name": "성명",
        "final_text": "최종 문장",
        "byte": "byte",
        "created_at": "생성일시",
        "generated_text": "생성 원문",
        "extra_comment": "추가 코멘트",
    }
    result_df = result_df.rename(columns=column_rename)
    preferred_cols = [
        "생성방식", "수행평가/관찰영역", "성취수준", "문장번호", "반", "번호", "성명", "최종 문장", "byte", "생성일시", "생성 원문", "추가 코멘트"
    ]
    for col in preferred_cols:
        if col not in result_df.columns:
            result_df[col] = ""
    result_df = result_df[preferred_cols]

    area_rows = []
    for area in st.session_state.mid_areas:
        for level in area.get("levels", []):
            area_rows.append(
                {
                    "수행평가/관찰영역": area.get("name", ""),
                    "단원/영역": area.get("unit", ""),
                    "활동 설명": area.get("description", ""),
                    "성취수준": level,
                    "평가 문구": area.get("rubrics", {}).get(level, ""),
                    "사용": area.get("use", True),
                    "순서": area.get("order", ""),
                }
            )
    area_df = pd.DataFrame(area_rows)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="중학교생기부문장", index=False)
        area_df.to_excel(writer, sheet_name="관찰영역_성취수준", index=False)

        wb = writer.book

        def style_ws(ws, header_fill_color="1E3A8A", tab_color="FF2563EB", freeze_cell="A2"):
            ws.sheet_view.showGridLines = False
            ws.sheet_properties.tabColor = tab_color
            ws.freeze_panes = freeze_cell
            header_fill = PatternFill("solid", fgColor=header_fill_color)
            header_font = Font(bold=True, color="FFFFFF", size=11)
            body_font = Font(size=10, color="111827")
            side = Side(style="thin", color="D1D5DB")
            border = Border(left=side, right=side, top=side, bottom=side)

            if ws.max_row >= 1:
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = border
                ws.row_dimensions[1].height = 28

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.font = body_font
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    cell.border = border
                ws.row_dimensions[row[0].row].height = 42

            if ws.max_row >= 1 and ws.max_column >= 1:
                ws.auto_filter.ref = ws.dimensions

            text_heavy = {"최종 문장", "생성 원문", "추가 코멘트", "활동 설명", "평가 문구"}
            for col_idx in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col_idx)
                header = clean_text(ws.cell(1, col_idx).value)
                if header in text_heavy:
                    width = 54 if header in ["최종 문장", "생성 원문"] else 34
                elif header in ["반", "번호", "byte", "순서"]:
                    width = 9
                elif header in ["성명", "성취수준", "문장번호"]:
                    width = 13
                else:
                    width = 22
                ws.column_dimensions[col_letter].width = width

        style_ws(wb["중학교생기부문장"], header_fill_color="1E3A8A", tab_color="FF2563EB", freeze_cell="H2")
        style_ws(wb["관찰영역_성취수준"], header_fill_color="92400E", tab_color="FFF59E0B", freeze_cell="A2")
        wb.active = wb.sheetnames.index("중학교생기부문장")

    output.seek(0)
    return output


# =========================
# 앱 UI 스타일
# =========================
st.markdown(
    """
    <style>
    /* 기존 app.py의 단계 탭 라디오 스타일과 맞춤 */
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

    /* 수행평가 박스는 app.py와 같은 연한 파란색 계열 */
    div[data-testid="stExpander"] details:has(.assessment-card-content) {
        background: linear-gradient(180deg, #EFF6FF 0%, #DBEAFE 100%) !important;
        border: 2px solid #93C5FD !important;
        border-radius: 18px !important;
        box-shadow: 0 5px 15px rgba(37, 99, 235, 0.10) !important;
        padding: 0.18rem 0.38rem 0.42rem 0.38rem !important;
        margin: 0.9rem 0 1.25rem 0 !important;
    }
    div[data-testid="stExpander"] details:has(.assessment-card-content) > summary {
        background: #DBEAFE !important;
        border: 1px solid #BFDBFE !important;
        border-radius: 14px !important;
        margin: 0.2rem 0 0.55rem 0 !important;
        padding: 0.15rem 0.45rem !important;
    }
    div[data-testid="stExpander"] details:has(.assessment-card-content) > summary p {
        color: #0F172A !important;
        font-weight: 900 !important;
    }
    div[data-testid="stExpander"] details:has(.assessment-card-content) .assessment-card-content {
        display: none !important;
    }
    div[data-testid="stMetric"] {
        background: #F9FAFB;
        border: 1px solid #E5E7EB;
        border-radius: 14px;
        padding: 0.55rem 0.75rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# 사이드바: app.py와 같은 작업 관리 위치
# =========================
with st.sidebar:
    st.header("작업 관리")

    uploaded_mid_project = st.file_uploader(
        "프로젝트 JSON 불러오기",
        type=["json"],
        help="중학교 간편 생성기에서 저장한 프로젝트 파일을 다시 불러옵니다.",
    )

    if uploaded_mid_project and st.button("프로젝트 불러오기"):
        load_mid_project_json(uploaded_mid_project)
        st.success("프로젝트를 불러왔습니다.")
        st.rerun()

    st.download_button(
        "현재 프로젝트 JSON 저장",
        data=mid_project_to_json(),
        file_name="개꿀생기부_mid_project.json",
        mime="application/json",
    )

    st.divider()

    if st.button("샘플 데이터 불러오기", help="중학교 간편 생성기용 가상 수행평가와 학생별 배정 예시를 불러옵니다."):
        apply_mid_project_data(build_mid_sample_project())
        st.success("샘플 데이터를 불러왔습니다.")
        st.rerun()

    if st.button("전체 초기화", type="secondary"):
        for key in ["mid_settings", "mid_areas", "mid_results", "mid_student_rows", "mid_selected_result_id"]:
            if key in st.session_state:
                del st.session_state[key]
        init_mid_state()
        sanitize_mid_state()
        st.success("초기화했습니다.")
        st.rerun()


# =========================
# 단계 이동: app.py 방식과 통일
# =========================
STEP_LABELS = [
    "① 기본 설정",
    "② 수행평가 설계",
    "③ 수준별 묶음 생성",
    "④ 학생별 배정 생성",
    "⑤ API 자료 확인",
    "⑥ 결과 수정/다운로드",
]

NAV_WIDGET_KEY = "mid_step_nav_radio_v02"
PENDING_STEP_KEY = "mid_pending_step_index_v02"
SCROLL_TO_TOP_KEY = "mid_scroll_to_top_after_step_change_v02"

if "mid_current_step" not in st.session_state:
    st.session_state["mid_current_step"] = 0

try:
    st.session_state["mid_current_step"] = int(st.session_state.get("mid_current_step", 0))
except Exception:
    st.session_state["mid_current_step"] = 0

programmatic_step_change = False
if PENDING_STEP_KEY in st.session_state:
    try:
        st.session_state["mid_current_step"] = int(st.session_state[PENDING_STEP_KEY])
    except Exception:
        st.session_state["mid_current_step"] = 0
    del st.session_state[PENDING_STEP_KEY]
    st.session_state[SCROLL_TO_TOP_KEY] = True
    programmatic_step_change = True

st.session_state["mid_current_step"] = max(0, min(st.session_state["mid_current_step"], len(STEP_LABELS) - 1))

if NAV_WIDGET_KEY not in st.session_state or programmatic_step_change:
    st.session_state[NAV_WIDGET_KEY] = STEP_LABELS[st.session_state["mid_current_step"]]


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
                const selectors = [
                    'section.main', 'main', '.main', '.block-container',
                    '[data-testid="stAppViewContainer"]', '[data-testid="stMain"]', '[data-testid="stMainBlockContainer"]'
                ];
                selectors.forEach(function(selector) {
                    parentDoc.querySelectorAll(selector).forEach(function(el) {
                        try { el.scrollTop = 0; } catch (e) {}
                        try { el.scrollTo(0, 0); } catch (e) {}
                    });
                });
                parentDoc.querySelectorAll('section, main, div').forEach(function(el) {
                    try {
                        if (el.scrollHeight > el.clientHeight + 80) {
                            el.scrollTop = 0;
                        }
                    } catch (e) {}
                });
            } catch (e) {}
        }
        forceScrollTop();
        setTimeout(forceScrollTop, 50);
        setTimeout(forceScrollTop, 150);
        setTimeout(forceScrollTop, 350);
        setTimeout(forceScrollTop, 700);
        setTimeout(forceScrollTop, 1100);
        </script>
        """,
        height=0,
    )


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
# 공통 렌더 함수
# =========================
def render_level_input_block(prefix, current_levels=None, current_rubrics=None):
    current_levels = current_levels if isinstance(current_levels, list) else []
    current_rubrics = current_rubrics if isinstance(current_rubrics, dict) else {}
    default_count = len(current_levels) if current_levels else 3
    default_count = max(1, min(10, int(default_count)))

    level_count = st.selectbox(
        "성취수준 개수",
        options=list(range(1, 11)),
        index=default_count - 1,
        key=f"{prefix}_level_count",
    )

    levels = []
    rubrics = {}
    st.markdown("**성취수준 코드와 성취 수준별 교사의 평가 문구**")
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
                placeholder="상",
            )
        with col_text:
            comment = st.text_area(
                f"{idx + 1}번 평가 문구",
                value=existing_text,
                key=f"{prefix}_rubric_{idx}",
                label_visibility="collapsed",
                height=76,
                placeholder="예: 활동의 핵심 내용을 바탕으로 결과를 설명함",
            )
        code = clean_text(code)
        if code:
            levels.append(code)
            rubrics[code] = clean_text(comment)
    return levels, rubrics


def render_ai_settings(prefix):
    st.markdown("#### AI 선택 및 API 설정")
    col_provider, col_model, col_key = st.columns([1.25, 1.7, 2.6], gap="small")
    with col_provider:
        provider = st.selectbox("사용할 AI", AI_PROVIDER_OPTIONS, index=0, key=f"{prefix}_provider")
    with col_model:
        model = st.text_input(
            "모델명",
            value=get_default_ai_model(provider),
            key=f"{prefix}_model_{provider}",
            help="필요하면 직접 수정할 수 있습니다.",
        )
    with col_key:
        api_key = st.text_input(
            f"{provider} API Key",
            value=get_default_ai_key(provider),
            type="password",
            key=f"{prefix}_api_key_{provider}",
            help=f"Streamlit Secrets에는 {AI_SECRET_KEY_NAMES.get(provider, 'OPENAI_API_KEY')} 이름으로 저장할 수 있습니다. 비워두면 API 없이 간단 변주 방식으로 생성됩니다.",
        )
    return provider, model, api_key


def get_selected_area(prefix):
    usable_areas = [area for area in sort_areas(st.session_state.mid_areas) if area.get("use", True)]
    if not usable_areas:
        return None
    options = [area_label(area, idx) for idx, area in enumerate(usable_areas, start=1)]
    selected_label = st.selectbox("수행평가/관찰 영역 선택", options, key=f"{prefix}_area_select")
    selected_index = options.index(selected_label)
    return usable_areas[selected_index]


# =========================
# 메인 화면
# =========================
st.markdown('<div id="mid-honey-top"></div>', unsafe_allow_html=True)
st.title(MID_APP_TITLE)
st.caption(MID_APP_SUBTITLE)

st.markdown("### 작업 단계")
selected_step_label = st.radio(
    "이동할 단계를 선택하세요.",
    STEP_LABELS,
    horizontal=True,
    label_visibility="collapsed",
    key=NAV_WIDGET_KEY,
)
st.session_state["mid_current_step"] = STEP_LABELS.index(selected_step_label)
current_step = st.session_state["mid_current_step"]
scroll_page_to_top_once()


# =========================
# ① 기본 설정
# =========================
if current_step == 0:
    st.subheader("① 기본 설정")
    settings = st.session_state.mid_settings

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        settings["school_year"] = st.text_input("학년도", value=clean_text(settings.get("school_year", "2026")), key="mid_school_year")
    with col2:
        semester_options = ["1학기", "2학기"]
        settings["semester"] = st.selectbox(
            "학기",
            semester_options,
            index=semester_options.index(settings.get("semester", "1학기")) if settings.get("semester") in semester_options else 0,
            key="mid_semester",
        )
    with col3:
        school_options = ["중학교", "고등학교", "초등학교"]
        current_school = settings.get("school_level", "중학교")
        settings["school_level"] = st.selectbox(
            "학교급",
            school_options,
            index=school_options.index(current_school) if current_school in school_options else 0,
            key="mid_school_level",
        )
    with col4:
        settings["grade"] = st.text_input("학년", value=clean_text(settings.get("grade", "2")), key="mid_grade")

    col5, col6, col7 = st.columns(3)
    with col5:
        settings["subject"] = st.text_input("과목명", value=clean_text(settings.get("subject", "과학")), key="mid_subject")
    with col6:
        settings["target_bytes_min"] = st.number_input(
            "목표 최소 byte",
            min_value=50,
            max_value=2000,
            value=int(settings.get("target_bytes_min", 250)),
            step=50,
            key="mid_target_min",
        )
    with col7:
        settings["target_bytes_max"] = st.number_input(
            "목표 최대 byte",
            min_value=50,
            max_value=2000,
            value=int(settings.get("target_bytes_max", 450)),
            step=50,
            key="mid_target_max",
        )

    st.markdown("#### 생기부 작성 규칙")
    st.caption("여기에 적은 규칙은 API 프롬프트의 공통 규칙으로 들어갑니다.")
    settings["custom_rules"] = st.text_area(
        "공통 작성 규칙",
        value=settings.get("custom_rules", MID_DEFAULT_RULES),
        height=220,
        key="mid_custom_rules",
    )

    with st.expander("API 프롬프트에 들어가는 규칙 예시 보기"):
        st.code(settings["custom_rules"], language="text")

    st.info("작업을 이어서 하려면 왼쪽의 '현재 프로젝트 JSON 저장'을 눌러 파일로 저장하세요.")
    render_next_step_button(0)


# =========================
# ② 수행평가 설계
# =========================
if current_step == 1:
    st.subheader("② 수행평가 설계")

    st.markdown(
        """
        중학교용도 **수행평가/관찰 영역**을 먼저 만들고, 각 영역 안에 **성취수준 코드와 평가 문구**를 입력합니다.  
        이후 ③ 또는 ④에서 같은 수준의 문장을 여러 개 변주하여 생성합니다.
        """
    )

    with st.expander("➕ 새 수행평가/관찰 영역 추가", expanded=True):
        st.caption("먼저 상위 단위인 수행평가 또는 관찰 영역을 만들고, 그 안에 성취수준을 설정합니다.")

        with st.form("mid_add_area_form"):
            col1, col2 = st.columns([2, 1])
            with col1:
                new_area_name = st.text_input("수행평가명", placeholder="예: 소화 기관 모형 만들기")
                new_unit = st.text_input("영역/단원", placeholder="예: 동물의 몸과 영양소")
            with col2:
                new_use = st.checkbox("사용", value=True)
                st.caption("순서는 아래의 드래그 정렬에서 바꿀 수 있습니다.")

            new_desc = st.text_area(
                "성취기준 / 활동 설명",
                placeholder="예: 소화 기관의 구조와 역할을 모형으로 표현하고 설명하는 활동",
                height=80,
            )

            submitted = st.form_submit_button("수행평가 추가")
            if submitted:
                if not clean_text(new_area_name):
                    st.warning("수행평가명을 입력하세요.")
                else:
                    st.session_state.mid_areas.append(
                        {
                            "area_id": make_id("mid_area"),
                            "name": clean_text(new_area_name),
                            "unit": clean_text(new_unit),
                            "description": clean_text(new_desc),
                            "order": len(st.session_state.mid_areas) + 1,
                            "use": new_use,
                            "levels": ["상", "중", "하"],
                            "rubrics": {
                                "상": "활동의 핵심 내용을 정확히 설명하고 결과를 구체적으로 정리함",
                                "중": "활동의 주요 내용을 바탕으로 결과를 대체로 설명함",
                                "하": "활동의 기본 내용을 중심으로 과제 수행에 참여함",
                            },
                        }
                    )
                    sanitize_mid_state()
                    st.success("수행평가를 추가했습니다.")
                    st.rerun()

    st.divider()
    st.markdown("### 📚 AI가 참고할 수행평가별 평가 자료")

    if not st.session_state.mid_areas:
        st.info("아직 등록된 수행평가가 없습니다. 먼저 수행평가를 추가하거나 왼쪽의 샘플 데이터를 불러오세요.")
    else:
        normalize_area_orders()
        sorted_areas = sort_areas(st.session_state.mid_areas)

        if len(sorted_areas) >= 2:
            with st.expander("수행평가 순서 드래그 정렬", expanded=True):
                if sort_items is None:
                    st.warning("드래그 정렬 기능을 사용하려면 requirements.txt에 streamlit-sortables가 있어야 합니다.")
                else:
                    st.caption("수행평가명을 마우스로 잡고 위아래로 옮긴 뒤 저장하세요.")
                    area_labels = [area_label(area, idx) for idx, area in enumerate(sorted_areas, start=1)]
                    st.caption("현재 수행평가 순서: " + " → ".join(area_labels))
                    label_to_area_id = {label: area.get("area_id", "") for label, area in zip(area_labels, sorted_areas)}
                    sorted_labels = sort_labels_with_gray_box(area_labels, key="mid_area_drag_sort", header="수행평가 순서")

                    if st.button("수행평가 순서 저장"):
                        apply_area_drag_order(sorted_labels, label_to_area_id)
                        normalize_area_orders()
                        st.success("수행평가 순서를 저장했습니다.")
                        st.rerun()

        level_updates = {}

        for area_index, area in enumerate(sorted_areas, start=1):
            aid = area.get("area_id", "")
            levels = [clean_text(x) for x in area.get("levels", []) if clean_text(x)]
            status_badge = "사용" if area.get("use", True) else "미사용"
            unit_text = area.get("unit", "") or "영역/단원 미입력"
            area_expander_title = f"📁 수행평가 {area_index}. {area.get('name', '이름 없는 수행평가')} · 성취수준 {len(levels)}개 · {status_badge}"

            with st.expander(area_expander_title, expanded=True):
                st.markdown('<div class="assessment-card-content"></div>', unsafe_allow_html=True)
                st.markdown(
                    f"""
                    ### 📁 수행평가 {area_index}. {area.get('name', '이름 없는 수행평가')}
                    **영역/단원:** {unit_text} &nbsp;&nbsp;|&nbsp;&nbsp;
                    **성취수준:** {len(levels)}개 &nbsp;&nbsp;|&nbsp;&nbsp;
                    **상태:** {status_badge}
                    """
                )

                if area.get("description"):
                    st.caption(f"활동 설명: {area.get('description', '')}")
                else:
                    st.caption("활동 설명이 아직 입력되지 않았습니다.")

                with st.expander("⚙️ 수행평가 기본 정보 수정 / 삭제", expanded=False):
                    col1, col2, col3 = st.columns([2, 2, 1])
                    with col1:
                        area["name"] = st.text_input("수행평가명 수정", value=area.get("name", ""), key=f"mid_area_name_{aid}")
                        area["unit"] = st.text_input("영역/단원 수정", value=area.get("unit", ""), key=f"mid_area_unit_{aid}")
                    with col2:
                        area["description"] = st.text_area("활동 설명 수정", value=area.get("description", ""), height=120, key=f"mid_area_desc_{aid}")
                    with col3:
                        st.caption("순서는 수행평가 목록 위의 드래그 정렬에서 변경합니다.")
                        area["use"] = st.checkbox("사용", value=area.get("use", True), key=f"mid_area_use_{aid}")
                        if st.button("수행평가 삭제", key=f"mid_delete_area_{aid}"):
                            st.session_state.mid_areas = [x for x in st.session_state.mid_areas if x.get("area_id", "") != aid]
                            st.session_state.mid_results = [x for x in st.session_state.mid_results if x.get("area_id", "") != aid]
                            st.success("수행평가를 삭제했습니다.")
                            st.rerun()

                st.markdown("#### 🧾 성취수준")
                new_levels, new_rubrics = render_level_input_block(
                    prefix=f"mid_area_levels_{aid}",
                    current_levels=area.get("levels", []),
                    current_rubrics=area.get("rubrics", {}),
                )
                level_updates[aid] = {"levels": new_levels, "rubrics": new_rubrics}
                st.caption("성취수준/평가 문구는 화면 맨 아래의 전체 저장 버튼으로 한꺼번에 저장됩니다.")

        if level_updates:
            st.divider()
            if st.button("전체 성취수준/평가 문구 한꺼번에 저장", type="primary", use_container_width=True):
                saved_count = 0
                for area in st.session_state.mid_areas:
                    aid = area.get("area_id", "")
                    if aid in level_updates:
                        area["levels"] = level_updates[aid]["levels"]
                        area["rubrics"] = level_updates[aid]["rubrics"]
                        saved_count += 1
                sanitize_mid_state()
                st.success(f"성취수준/평가 문구를 {saved_count}개 수행평가에 한꺼번에 저장했습니다.")
                st.rerun()

    render_next_step_button(1)


# =========================
# ③ 수준별 묶음 생성
# =========================
if current_step == 2:
    st.subheader("③ 수준별 문장 묶음 생성")
    st.caption("학생 이름 없이, 성취수준별로 여러 개의 문장을 한 번에 만듭니다. 나중에 필요한 학생에게 골라 붙이기 좋습니다.")

    if not st.session_state.mid_areas:
        st.warning("먼저 ②에서 수행평가/관찰 영역을 추가하세요.")
    else:
        selected_area = get_selected_area("mid_batch")
        if selected_area is None:
            st.warning("사용 설정된 수행평가가 없습니다.")
        else:
            levels = [clean_text(x) for x in selected_area.get("levels", []) if clean_text(x)]
            st.markdown("#### 수준별 생성 개수")
            level_counts = {}
            cols = st.columns(min(5, max(1, len(levels))))
            for idx, level in enumerate(levels):
                with cols[idx % len(cols)]:
                    level_counts[level] = st.number_input(
                        f"{level} 수준",
                        min_value=0,
                        max_value=50,
                        value=3,
                        step=1,
                        key=f"mid_batch_count_{selected_area.get('area_id')}_{level}",
                    )

            variation_strength = st.radio(
                "변주 강도",
                ["낮음", "보통", "높음"],
                index=1,
                horizontal=True,
                key="mid_batch_variation_strength",
                help="낮음은 복붙 느낌을 조금 줄이는 정도, 높음은 문장 구조를 더 다양하게 바꿉니다.",
            )

            provider, model, api_key = render_ai_settings("mid_batch_ai")

            if st.button("수준별 문장 묶음 생성", type="primary", use_container_width=True):
                total_requested = sum(int(v) for v in level_counts.values())
                if total_requested <= 0:
                    st.warning("생성할 문장 개수를 1개 이상 입력하세요.")
                else:
                    new_results = []
                    with st.spinner("수준별 문장 변주를 생성하는 중..."):
                        for level, count in level_counts.items():
                            if int(count) <= 0:
                                continue
                            sentences, prompt = generate_variations(selected_area, level, int(count), provider, api_key, model, variation_strength)
                            for seq, sentence in enumerate(sentences, start=1):
                                new_results.append(
                                    {
                                        "result_id": make_id("mid_result"),
                                        "mode": "수준별 묶음",
                                        "area_id": selected_area.get("area_id", ""),
                                        "area_name": selected_area.get("name", ""),
                                        "level": level,
                                        "variant_no": seq,
                                        "class": "",
                                        "number": "",
                                        "name": "",
                                        "generated_text": sentence,
                                        "final_text": sentence,
                                        "byte": byte_count(sentence),
                                        "extra_comment": "",
                                        "prompt": prompt,
                                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    }
                                )
                    st.session_state.mid_results.extend(new_results)
                    st.success(f"문장 {len(new_results)}개를 생성했습니다. ⑥ 결과 수정/다운로드에서 확인하세요.")
                    st.rerun()

    render_next_step_button(2)


# =========================
# ④ 학생별 배정 생성
# =========================
if current_step == 3:
    st.subheader("④ 학생별 배정 생성")
    st.caption("학생별로 성취수준과 간단한 추가 코멘트를 넣으면, 각 학생에게 붙일 수 있는 문장을 생성합니다. 학생 이름은 AI 입력 자료에 넣지 않습니다.")

    if not st.session_state.mid_areas:
        st.warning("먼저 ②에서 수행평가/관찰 영역을 추가하세요.")
    else:
        selected_area = get_selected_area("mid_student")
        if selected_area is None:
            st.warning("사용 설정된 수행평가가 없습니다.")
        else:
            levels = [clean_text(x) for x in selected_area.get("levels", []) if clean_text(x)]

            if st.session_state.mid_student_rows.empty:
                st.session_state.mid_student_rows = pd.DataFrame(
                    [{"반": "1", "번호": "1", "성명": "", "성취수준": levels[0] if levels else "", "추가 코멘트": "", "생성 수": 1}]
                )

            edited_student_rows = st.data_editor(
                st.session_state.mid_student_rows,
                use_container_width=True,
                num_rows="dynamic",
                height=360,
                key=f"mid_student_rows_editor_{selected_area.get('area_id')}",
                column_config={
                    "반": st.column_config.TextColumn("반", width="small"),
                    "번호": st.column_config.TextColumn("번호", width="small"),
                    "성명": st.column_config.TextColumn("성명", width="medium"),
                    "성취수준": st.column_config.SelectboxColumn("성취수준", options=[""] + levels, width="small"),
                    "추가 코멘트": st.column_config.TextColumn("추가 코멘트", width="large"),
                    "생성 수": st.column_config.NumberColumn("생성 수", min_value=1, max_value=5, step=1, width="small"),
                },
            )

            col_save_rows, col_hint = st.columns([1.5, 4])
            with col_save_rows:
                if st.button("학생별 배정표 저장", type="secondary", use_container_width=True):
                    st.session_state.mid_student_rows = edited_student_rows.copy()
                    st.success("학생별 배정표를 저장했습니다.")
                    st.rerun()
            with col_hint:
                st.caption("성명은 결과표와 엑셀에만 표시됩니다. AI 프롬프트에는 이름, 반, 번호를 넣지 않습니다.")

            variation_strength = st.radio("변주 강도", ["낮음", "보통", "높음"], index=1, horizontal=True, key="mid_student_variation_strength")
            provider, model, api_key = render_ai_settings("mid_student_ai")

            if st.button("학생별 문장 생성", type="primary", use_container_width=True):
                st.session_state.mid_student_rows = edited_student_rows.copy()
                new_results = []
                usable_rows = edited_student_rows.copy()
                usable_rows = usable_rows[usable_rows["성취수준"].map(clean_text) != ""].reset_index(drop=True)

                if usable_rows.empty:
                    st.warning("성취수준이 입력된 학생 행이 없습니다.")
                else:
                    with st.spinner("학생별 문장을 생성하는 중..."):
                        for _, row in usable_rows.iterrows():
                            level = clean_text(row.get("성취수준", ""))
                            if level not in levels:
                                continue
                            try:
                                gen_count = int(row.get("생성 수", 1) or 1)
                            except Exception:
                                gen_count = 1
                            gen_count = max(1, min(5, gen_count))
                            extra_comment = clean_text(row.get("추가 코멘트", ""))
                            sentences, prompt = generate_variations(
                                selected_area, level, gen_count, provider, api_key, model, variation_strength, extra_comment=extra_comment
                            )
                            for seq, sentence in enumerate(sentences, start=1):
                                new_results.append(
                                    {
                                        "result_id": make_id("mid_result"),
                                        "mode": "학생별 배정",
                                        "area_id": selected_area.get("area_id", ""),
                                        "area_name": selected_area.get("name", ""),
                                        "level": level,
                                        "variant_no": seq,
                                        "class": clean_text(row.get("반", "")),
                                        "number": clean_text(row.get("번호", "")),
                                        "name": clean_text(row.get("성명", "")),
                                        "generated_text": sentence,
                                        "final_text": sentence,
                                        "byte": byte_count(sentence),
                                        "extra_comment": extra_comment,
                                        "prompt": prompt,
                                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    }
                                )
                    st.session_state.mid_results.extend(new_results)
                    st.success(f"학생별 문장 {len(new_results)}개를 생성했습니다. ⑥ 결과 수정/다운로드에서 확인하세요.")
                    st.rerun()

    render_next_step_button(3)


# =========================
# ⑤ API 자료 확인
# =========================
if current_step == 4:
    st.subheader("⑤ API 입력 자료 확인")

    if not st.session_state.mid_areas:
        st.warning("먼저 수행평가/관찰 영역을 입력하세요.")
    else:
        selected_area = get_selected_area("mid_api_preview")
        if selected_area is None:
            st.warning("사용 설정된 수행평가가 없습니다.")
        else:
            levels = [clean_text(x) for x in selected_area.get("levels", []) if clean_text(x)]
            col_level, col_count, col_variation = st.columns([1, 1, 1.2])
            with col_level:
                selected_level = st.selectbox("성취수준 선택", levels, key="mid_api_preview_level")
            with col_count:
                preview_count = st.number_input("생성 개수", min_value=1, max_value=10, value=3, step=1, key="mid_api_preview_count")
            with col_variation:
                preview_variation = st.selectbox("변주 강도", ["낮음", "보통", "높음"], index=1, key="mid_api_preview_variation")

            preview_extra = st.text_input("추가 코멘트 예시", value="", key="mid_api_preview_extra")
            rubric = selected_area.get("rubrics", {}).get(selected_level, "")
            preview_prompt = build_mid_generation_prompt(
                selected_area,
                selected_level,
                rubric,
                int(preview_count),
                preview_variation,
                extra_comment=preview_extra,
            )

            st.markdown("#### API 입력 자료 미리보기")
            material = f"""- 수행평가/관찰 영역: {selected_area.get('name', '')}
- 영역/단원: {selected_area.get('unit', '')}
- 활동 설명: {selected_area.get('description', '')}
- 성취수준: 전체 성취수준 {', '.join(levels)} 중 {selected_level}
- 교사의 평가: {rubric}"""
            if preview_extra:
                material += f"\n- 추가 코멘트: {preview_extra}"
            st.text_area("이 내용이 API 프롬프트의 평가 자료로 들어갑니다.", value=material, height=220)

            with st.expander("실제 프롬프트 보기", expanded=True):
                st.text_area("프롬프트", value=preview_prompt, height=420)

    render_next_step_button(4)


# =========================
# ⑥ 결과 수정/다운로드
# =========================
if current_step == 5:
    st.subheader("⑥ 결과 수정 / 다운로드")

    if not st.session_state.mid_results:
        st.info("아직 생성된 문장이 없습니다. ③ 또는 ④에서 먼저 문장을 생성하세요.")
    else:
        results_df = pd.DataFrame(st.session_state.mid_results)
        display_cols = [
            "result_id", "mode", "area_name", "level", "variant_no", "class", "number", "name", "final_text", "byte", "created_at", "generated_text", "extra_comment", "prompt"
        ]
        for col in display_cols:
            if col not in results_df.columns:
                results_df[col] = ""
        results_df["byte"] = results_df["final_text"].map(byte_count)
        results_df = results_df[display_cols]

        valid_result_ids = set(results_df["result_id"].astype(str).tolist())
        if clean_text(st.session_state.get("mid_selected_result_id", "")) not in valid_result_ids:
            st.session_state.mid_selected_result_id = clean_text(results_df.iloc[0].get("result_id", ""))

        st.markdown("#### 전체 결과표 / 문장 선택")
        st.caption("표에서 아무 셀이나 클릭하면 해당 문장이 아래 큰 수정창의 대상으로 선택됩니다. 직접 수정은 아래 큰 수정창에서 처리하세요.")

        display_result_df = results_df.drop(columns=["result_id", "generated_text", "extra_comment", "prompt"], errors="ignore")
        selection_event = st.dataframe(
            display_result_df,
            use_container_width=True,
            height=360,
            hide_index=True,
            key="mid_result_selector",
            on_select="rerun",
            selection_mode="single-cell",
            column_config={
                "mode": st.column_config.TextColumn("생성방식", width="medium"),
                "area_name": st.column_config.TextColumn("수행평가/관찰영역", width="medium"),
                "level": st.column_config.TextColumn("성취수준", width="small"),
                "variant_no": st.column_config.NumberColumn("문장번호", width="small"),
                "class": st.column_config.TextColumn("반", width="small"),
                "number": st.column_config.TextColumn("번호", width="small"),
                "name": st.column_config.TextColumn("성명", width="medium"),
                "final_text": st.column_config.TextColumn("생성/수정 문구", width="large"),
                "byte": st.column_config.NumberColumn("byte", width="small"),
                "created_at": st.column_config.TextColumn("생성일시", width="medium"),
            },
        )

        selected_row_index = None
        try:
            selected_cells = list(selection_event.selection.cells)
        except Exception:
            try:
                selected_cells = list(selection_event.get("selection", {}).get("cells", []))
            except Exception:
                selected_cells = []
        if selected_cells:
            first_cell = selected_cells[-1]
            try:
                selected_row_index = int(first_cell[0])
            except Exception:
                selected_row_index = None

        if selected_row_index is not None and 0 <= selected_row_index < len(results_df):
            st.session_state.mid_selected_result_id = clean_text(results_df.iloc[selected_row_index].get("result_id", ""))

        with st.expander("표 형태로 여러 문구 한꺼번에 수정", expanded=False):
            st.caption("여러 문구를 표에서 한꺼번에 고칠 때만 사용하세요. 문장 선택은 위 표의 아무 셀이나 클릭해서 합니다.")
            edited_results = st.data_editor(
                results_df,
                use_container_width=True,
                height=360,
                hide_index=True,
                key="mid_results_bulk_editor",
                disabled=["result_id", "mode", "area_name", "level", "variant_no", "class", "number", "name", "byte", "created_at", "generated_text", "extra_comment", "prompt"],
                column_config={
                    "result_id": None,
                    "mode": st.column_config.TextColumn("생성방식", width="medium"),
                    "area_name": st.column_config.TextColumn("수행평가/관찰영역", width="medium"),
                    "level": st.column_config.TextColumn("성취수준", width="small"),
                    "variant_no": st.column_config.NumberColumn("문장번호", width="small"),
                    "class": st.column_config.TextColumn("반", width="small"),
                    "number": st.column_config.TextColumn("번호", width="small"),
                    "name": st.column_config.TextColumn("성명", width="medium"),
                    "final_text": st.column_config.TextColumn("생성/수정 문구", width="large"),
                    "byte": st.column_config.NumberColumn("byte", width="small"),
                    "created_at": st.column_config.TextColumn("생성일시", width="medium"),
                    "generated_text": st.column_config.TextColumn("생성 원문", width="large"),
                    "extra_comment": st.column_config.TextColumn("추가 코멘트", width="large"),
                    "prompt": None,
                },
            )

            col_save_table, col_clear_table, col_table_info = st.columns([1.5, 1.5, 4])
            with col_save_table:
                if st.button("표에서 수정한 문구 저장", type="primary", use_container_width=True):
                    new_results = []
                    for _, row in edited_results.iterrows():
                        row_dict = row.to_dict()
                        row_dict["final_text"] = clean_text(row_dict.get("final_text", ""))
                        row_dict["byte"] = byte_count(row_dict.get("final_text", ""))
                        new_results.append(row_dict)
                    st.session_state.mid_results = new_results
                    st.success(f"표에서 수정한 문구 {len(new_results)}건을 저장했습니다.")
                    st.rerun()
            with col_clear_table:
                if st.button("생성 결과 전체 삭제", type="secondary", use_container_width=True):
                    st.session_state.mid_results = []
                    st.warning("생성 결과를 모두 삭제했습니다.")
                    st.rerun()
            with col_table_info:
                st.caption("표의 byte는 저장 후 다시 계산되어 반영됩니다. 긴 문장은 아래 큰 수정창에서 고치면 현재 byte를 바로 확인하기 쉽습니다.")

        st.divider()
        selected_result_id = clean_text(st.session_state.get("mid_selected_result_id", ""))
        selected_result = None
        for result in st.session_state.mid_results:
            if clean_text(result.get("result_id", "")) == selected_result_id:
                selected_result = result
                break
        if selected_result is None:
            selected_result = st.session_state.mid_results[0]
            selected_result_id = clean_text(selected_result.get("result_id", ""))

        selected_label_parts = [
            selected_result.get("mode", ""),
            selected_result.get("area_name", ""),
            selected_result.get("level", ""),
        ]
        if selected_result.get("name"):
            selected_label_parts.append(f"{selected_result.get('class', '')}반 {selected_result.get('number', '')}번 {selected_result.get('name', '')}")
        selected_detail_label = " / ".join([clean_text(x) for x in selected_label_parts if clean_text(x)])

        st.markdown(f"#### 표에서 선택한 문장 수정: {selected_detail_label}")
        editor_key = f"mid_generation_detail_text_{selected_result_id}"
        if editor_key not in st.session_state:
            st.session_state[editor_key] = selected_result.get("final_text", selected_result.get("generated_text", ""))

        edited = st.text_area("교사 수정 문구", key=editor_key, height=220, help="입력한 내용의 byte를 아래에서 바로 확인할 수 있습니다.")
        current_bytes = byte_count(edited)
        st.metric("현재 byte", current_bytes)

        if st.button("수정 문구 저장", type="primary"):
            for idx, result in enumerate(st.session_state.mid_results):
                if clean_text(result.get("result_id", "")) == selected_result_id:
                    result["final_text"] = edited
                    result["byte"] = current_bytes
                    if not result.get("generated_text"):
                        result["generated_text"] = edited
                    st.session_state.mid_results[idx] = result
                    break
            st.success("수정 문구를 저장했습니다.")
            st.rerun()

        with st.expander("선택한 결과의 API 입력 프롬프트 확인", expanded=False):
            st.text_area("API 입력 프롬프트", value=selected_result.get("prompt", ""), height=360)

        st.divider()
        current_bytes_list = [byte_count(x.get("final_text", "")) for x in st.session_state.mid_results]
        col_metric1, col_metric2, col_metric3 = st.columns(3)
        with col_metric1:
            st.metric("생성 문장 수", len(st.session_state.mid_results))
        with col_metric2:
            st.metric("평균 byte", int(sum(current_bytes_list) / len(current_bytes_list)) if current_bytes_list else 0)
        with col_metric3:
            st.metric("최대 byte", max(current_bytes_list) if current_bytes_list else 0)

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
        excel_file = export_mid_excel()
        st.download_button(
            "📥 중학교 문장 결과 엑셀 다운로드",
            data=excel_file,
            file_name=f"개꿀생기부_mid_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
