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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =========================
# 중학교 간편 생성기 기본 설정
# =========================
st.set_page_config(
    page_title="중학교 간편 생기부",
    page_icon="🍯",
    layout="wide",
)

MID_APP_TITLE = "🍯 중학교 간편 생기부 문장 생성 v17"
MID_APP_SUBTITLE = "활동·관찰 영역과 성취수준을 바탕으로 비슷하지만 조금씩 다른 문장을 생성합니다 · patched-20260623-v17-mid-sample"
MID_APP_VERSION = "patched-20260623-v17-mid-sample"

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
    name = clean_text(area.get("name", "이름 없는 관찰 영역")) or "이름 없는 관찰 영역"
    unit = clean_text(area.get("unit", ""))
    return f"{prefix}{name}" + (f" · {unit}" if unit else "")


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

    # 관찰 영역별 학생 입력표를 따로 저장한다.
    # 기존 버전의 mid_student_rows는 첫 영역/현재 영역과의 호환용으로 유지한다.
    if "mid_student_rows_by_area" not in st.session_state:
        st.session_state.mid_student_rows_by_area = {}


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
        area["name"] = clean_text(area.get("name", "")) or "이름 없는 관찰 영역"
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
        for level in area["levels"]:
            area["rubrics"].setdefault(level, "")
        clean_areas.append(area)

    clean_areas = sort_areas(clean_areas)
    for idx, area in enumerate(clean_areas, start=1):
        area["order"] = idx
    st.session_state.mid_areas = clean_areas

    if not isinstance(st.session_state.get("mid_results"), list):
        st.session_state.mid_results = []

    if not isinstance(st.session_state.get("mid_student_rows"), pd.DataFrame):
        st.session_state.mid_student_rows = pd.DataFrame(st.session_state.get("mid_student_rows", []))
    for col in ["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]:
        if col not in st.session_state.mid_student_rows.columns:
            st.session_state.mid_student_rows[col] = "" if col != "생성 수" else 1

    if not isinstance(st.session_state.get("mid_student_rows_by_area"), dict):
        st.session_state.mid_student_rows_by_area = {}
    normalized_by_area = {}
    for area_id, rows in st.session_state.mid_student_rows_by_area.items():
        if isinstance(rows, pd.DataFrame):
            df = rows.copy()
        else:
            df = pd.DataFrame(rows or [])
        for col in ["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]:
            if col not in df.columns:
                df[col] = "" if col != "생성 수" else 1
        normalized_by_area[str(area_id)] = df[["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]].to_dict(orient="records")
    st.session_state.mid_student_rows_by_area = normalized_by_area


def ensure_student_rows_df(rows=None):
    if isinstance(rows, pd.DataFrame):
        df = rows.copy()
    else:
        df = pd.DataFrame(rows or [])
    for col in ["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]:
        if col not in df.columns:
            df[col] = "" if col != "생성 수" else 1
    return df[["반", "번호", "성명", "성취수준", "추가 코멘트", "생성 수"]].copy()


def get_student_rows_for_area(area_id: str, levels=None) -> pd.DataFrame:
    area_id = clean_text(area_id)
    by_area = st.session_state.get("mid_student_rows_by_area", {})
    if area_id and area_id in by_area:
        return ensure_student_rows_df(by_area.get(area_id))
    base_df = ensure_student_rows_df(st.session_state.get("mid_student_rows", []))
    if base_df.empty:
        default_level = clean_text(levels[0]) if levels else ""
        base_df = pd.DataFrame([{"반": "1", "번호": "1", "성명": "", "성취수준": default_level, "추가 코멘트": "", "생성 수": 1}])
    return ensure_student_rows_df(base_df)


def set_student_rows_for_area(area_id: str, df) -> None:
    area_id = clean_text(area_id)
    clean_df = ensure_student_rows_df(df)
    if area_id:
        st.session_state.mid_student_rows_by_area[area_id] = clean_df.to_dict(orient="records")
    st.session_state.mid_student_rows = clean_df.copy()


init_mid_state()
sanitize_mid_state()


# =========================
# JSON 저장/불러오기 및 샘플
# =========================
def mid_project_to_json() -> str:
    data = {
        "settings": st.session_state.mid_settings,
        "areas": st.session_state.mid_areas,
        "student_rows": ensure_student_rows_df(st.session_state.mid_student_rows).to_dict(orient="records"),
        "student_rows_by_area": {
            str(area_id): ensure_student_rows_df(rows).to_dict(orient="records")
            for area_id, rows in st.session_state.get("mid_student_rows_by_area", {}).items()
        },
        "results": st.session_state.mid_results,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "개꿀 생기부 - 중학교 간편 생성기",
        "version": MID_APP_VERSION,
    }
    return json.dumps(json_safe(data), ensure_ascii=False, indent=2, default=str)


def apply_mid_project_data(data):
    st.session_state.mid_settings = data.get("settings", st.session_state.mid_settings)
    st.session_state.mid_areas = data.get("areas", [])
    st.session_state.mid_student_rows = ensure_student_rows_df(data.get("student_rows", []))
    st.session_state.mid_student_rows_by_area = data.get("student_rows_by_area", {}) or {}
    if not st.session_state.mid_student_rows_by_area and not st.session_state.mid_student_rows.empty:
        first_area_id = clean_text(st.session_state.mid_areas[0].get("area_id", "")) if st.session_state.mid_areas else ""
        if first_area_id:
            st.session_state.mid_student_rows_by_area[first_area_id] = st.session_state.mid_student_rows.to_dict(orient="records")
    st.session_state.mid_results = data.get("results", [])
    sanitize_mid_state()


def load_mid_project_json(uploaded_file):
    data = json.load(uploaded_file)
    apply_mid_project_data(data)


def build_mid_sample_project():
    """사진 속 채점 기준을 바탕으로 만든 중학교 과학 샘플 프로젝트."""
    students = [
        ("1", "1", "김민준"), ("1", "2", "박서연"), ("1", "3", "최도윤"), ("1", "4", "이하은"), ("1", "5", "정우진"),
        ("1", "6", "한지우"), ("1", "7", "오서준"), ("1", "8", "윤채아"), ("1", "9", "서지호"), ("1", "10", "임하준"),
        ("2", "1", "강민서"), ("2", "2", "송예린"), ("2", "3", "장도현"), ("2", "4", "유나경"), ("2", "5", "조현우"),
        ("2", "6", "신아린"), ("2", "7", "배준서"), ("2", "8", "문태오"), ("2", "9", "서지민"), ("2", "10", "이도겸"),
        ("3", "1", "홍시우"), ("3", "2", "권예준"), ("3", "3", "남서아"), ("3", "4", "구민재"), ("3", "5", "하은별"),
        ("3", "6", "백다온"), ("3", "7", "민서율"), ("3", "8", "노하윤"), ("3", "9", "차지호"), ("3", "10", "유가온"),
    ]

    areas = [
        {
            "area_id": "mid_area_mix_sep_method",
            "name": "혼합물 분리하기",
            "unit": "혼합물 분리 탐구",
            "description": "혼합물의 성분을 파악하고 탐구 수행 단계별로 적절한 분리 방법을 선택하여 물질을 분리하는 활동",
            "order": 1,
            "use": True,
            "levels": ["6점", "5점", "4점", "3점", "0점"],
            "rubrics": {
                "6점": "탐구 수행 단계별로 가장 효율적인 분리 방법을 선택해 수행하여 물질을 모두 분리함",
                "5점": "탐구 수행 단계별로 분리 방법을 선택해 수행하여 물질을 모두 분리함",
                "4점": "선택한 분리 방법과 절차에 따라 수행했으나 일부 물질만 분리함",
                "3점": "선택한 분리 방법과 절차 중 일부만 수행하여 일부 물질만 분리함",
                "0점": "혼합물의 분리를 수행하지 못함",
            },
        },
        {
            "area_id": "mid_area_mix_property_reason",
            "name": "물질의 특성을 혼합물 분리 과정과 연결 짓기",
            "unit": "혼합물 분리 탐구",
            "description": "밀도, 용해도, 끓는점, 녹는점 등 물질의 특성을 혼합물 분리에 적용한 이유를 설명하는 활동",
            "order": 2,
            "use": True,
            "levels": ["7점", "5점", "3점", "0점"],
            "rubrics": {
                "7점": "단계별로 과학적인 근거를 들어 오류 없이 설명함",
                "5점": "모든 단계에 설명하였으나 일부 오류가 있음",
                "3점": "일부 단계만 설명함",
                "0점": "설명하지 못함",
            },
        },
        {
            "area_id": "mid_area_mix_report",
            "name": "과학 탐구 보고서 작성하기",
            "unit": "혼합물 분리 탐구",
            "description": "혼합물 분리 실험 계획서와 결과 보고서를 작성하는 활동",
            "order": 3,
            "use": True,
            "levels": ["4점", "2점", "0점"],
            "rubrics": {
                "4점": "실험 계획서와 결과 보고서를 빠진 내용이나 오류 없이 작성함",
                "2점": "실험 계획서와 결과 보고서가 일부 누락되었거나 일부 오류가 있음",
                "0점": "혼합물의 분리 실험 계획서와 결과 보고서를 작성하지 못함",
            },
        },
        {
            "area_id": "mid_area_mix_improvement",
            "name": "실험 과정의 한계나 오차 요인 찾고 개선하기",
            "unit": "혼합물 분리 탐구",
            "description": "실험의 오류와 한계를 분석하고 실현 가능한 개선 아이디어를 제안하는 활동",
            "order": 4,
            "use": True,
            "levels": ["3점", "2점", "0점"],
            "rubrics": {
                "3점": "실험의 오류와 한계를 분석하여 실현이 가능한 아이디어를 제안함",
                "2점": "실험의 오류와 한계를 분석하였으나 아이디어가 미흡함",
                "0점": "실험의 오류와 한계를 분석하지 못함",
            },
        },
        {
            "area_id": "mid_area_mix_comment",
            "name": "총평",
            "unit": "혼합물 분리 탐구",
            "description": "혼합물 분리 탐구에서 드러난 태도, 협업, 관찰 특징을 개별 코멘트로 입력하는 항목",
            "order": 5,
            "use": True,
            "levels": ["총평"],
            "rubrics": {"총평": "개별 코멘트를 바탕으로 혼합물 분리 탐구 활동의 총평을 작성함"},
        },
        {
            "area_id": "mid_area_photo_variable",
            "name": "환경요인과 광합성의 관계를 탐구하기 위한 적절한 변인 통제",
            "unit": "광합성과 환경 요인 탐구",
            "description": "광합성에 영향을 주는 환경요인 중 다르게 해야 할 조건과 같게 해야 할 조건을 구분하는 활동",
            "order": 6,
            "use": True,
            "levels": ["우수", "보통", "기초", "미흡"],
            "rubrics": {
                "우수": "광합성과 관계된 환경요인 중 다르게 해야 할 조건 1개와 같게 해야 할 조건 3가지를 명확히 서술했음",
                "보통": "광합성과 관계된 환경요인 중 다르게 해야 할 조건 1개와 같게 해야 할 조건 2가지만 명확히 서술했음",
                "기초": "광합성과 관계된 환경요인 중 다르게 해야 할 조건 1개와 같게 해야 할 조건 1가지만 명확히 서술했음",
                "미흡": "광합성과 관계된 환경요인 중 다르게 해야 할 조건 2개와 같게 해야 할 조건을 구분하지 못함",
            },
        },
        {
            "area_id": "mid_area_photo_process",
            "name": "실험 과정",
            "unit": "광합성과 환경 요인 탐구",
            "description": "주어진 가설, 측정 방법, 준비물을 보고 조건에 맞게 실제 실험 과정을 설계하는 활동",
            "order": 7,
            "use": True,
            "levels": ["우수", "보통", "기초", "미흡"],
            "rubrics": {
                "우수": "주어진 가설과 측정 방법, 준비물을 보고 조건에 맞추어 실제 실험을 할 수 있을 정도로 잘 설계함",
                "보통": "주어진 가설과 측정 방법, 준비물을 보고 조건에 맞추어 대체로 잘 설계함",
                "기초": "주어진 가설과 측정 방법, 준비물을 보고 조건에 맞추어 설계하였으나 부족한 부분이 있음",
                "미흡": "주어진 가설과 측정 방법, 준비물을 보고 조건에 맞추어 설계하지 못함",
            },
        },
        {
            "area_id": "mid_area_photo_prediction",
            "name": "자료 해석 및 결과 예측",
            "unit": "광합성과 환경 요인 탐구",
            "description": "변인에 따른 광합성 정도의 결과를 자료를 바탕으로 해석하고 예측하는 활동",
            "order": 8,
            "use": True,
            "levels": ["우수", "보통", "기초", "미흡"],
            "rubrics": {
                "우수": "변인에 따른 광합성 정도에 대한 결과를 자료를 바탕으로 올바르게 예측함",
                "보통": "변인에 따른 광합성 정도에 대한 결과를 자료를 바탕으로 대체로 잘 예측함",
                "기초": "변인에 따른 광합성 정도에 대한 결과를 자료를 바탕으로 예측을 시도했으나 부족한 부분이 있음",
                "미흡": "변인에 따른 광합성 정도에 대한 결과를 자료를 바탕으로 예측하지 못함",
            },
        },
        {
            "area_id": "mid_area_photo_comment",
            "name": "총평",
            "unit": "광합성과 환경 요인 탐구",
            "description": "광합성과 환경 요인 탐구에서 드러난 태도, 설계 습관, 자료 해석 특징을 개별 코멘트로 입력하는 항목",
            "order": 9,
            "use": True,
            "levels": ["총평"],
            "rubrics": {"총평": "개별 코멘트를 바탕으로 광합성과 환경 요인 탐구 활동의 총평을 작성함"},
        },
    ]

    level_patterns = {
        "mid_area_mix_sep_method": ["6점", "5점", "4점", "6점", "3점", "5점", "4점", "6점", "5점", "4점"],
        "mid_area_mix_property_reason": ["7점", "5점", "3점", "7점", "5점", "5점", "3점", "7점", "5점", "3점"],
        "mid_area_mix_report": ["4점", "4점", "2점", "4점", "2점", "4점", "2점", "4점", "4점", "2점"],
        "mid_area_mix_improvement": ["3점", "2점", "2점", "3점", "0점", "2점", "3점", "3점", "2점", "2점"],
        "mid_area_photo_variable": ["우수", "보통", "기초", "우수", "보통", "기초", "보통", "우수", "기초", "미흡"],
        "mid_area_photo_process": ["우수", "보통", "보통", "우수", "기초", "보통", "기초", "우수", "보통", "미흡"],
        "mid_area_photo_prediction": ["우수", "보통", "기초", "우수", "보통", "기초", "보통", "우수", "기초", "미흡"],
    }

    mix_comments = {
        0: "분리 순서를 먼저 그림으로 정리한 뒤 실험에 들어가서 모둠이 덜 헤매게 도와줌. 실험 도구를 조심스럽게 다루는 편임.",
        2: "커피 여과 장면을 예로 들어 여과 원리를 설명하려고 해서 친구들이 이해하기 쉬웠음.",
        4: "처음에는 분리 방법 선택을 어려워했지만, 교사의 힌트 뒤에 기준을 다시 세워 끝까지 참여함.",
        6: "오차가 생긴 이유를 단순 실수로 넘기지 않고 입자 크기와 거름종이 상태까지 살펴봄.",
        8: "보고서 정리는 느린 편이지만, 실험 중 관찰한 장면을 꽤 구체적으로 기억해서 적음.",
        10: "모둠원이 놓친 절차를 조용히 챙겨 주는 모습이 있었고, 분리 결과를 사진으로 남기자고 제안함.",
        12: "분리 방법을 한 번에 고르기보다 왜 그 방법이 맞는지 계속 확인하려는 태도가 보임.",
        14: "끓는점과 용해도 차이를 헷갈려 했으나, 예시 물질을 다시 대입하면서 스스로 수정함.",
        16: "실험 후 정리 시간에 책상을 끝까지 정리하고 남은 시료를 확인하는 등 책임감 있는 모습을 보임.",
        18: "결과가 예상과 다르게 나오자 바로 실패로 보지 않고 어떤 과정에서 달라졌는지 묻는 태도가 좋았음.",
        20: "모둠 발표 때 말수는 많지 않았지만, 분리 순서를 묻는 질문에는 핵심만 짧게 잘 대답함.",
        22: "활동 초반에는 소극적이었으나 직접 거름 장치를 잡아 본 뒤 참여도가 눈에 띄게 높아짐.",
        24: "혼합물 분리 결과를 색과 상태 변화 중심으로 설명해 관찰 내용이 잘 드러남.",
        26: "친구의 의견을 바로 부정하기보다 실험 절차에 맞춰 다시 비교해 보자고 말하는 모습이 인상적임.",
        28: "보고서의 표현은 간단했지만 실험에서 실제로 본 장면을 바탕으로 작성하려는 점이 좋았음.",
    }
    photo_comments = {
        1: "빛의 세기를 바꾸는 조건을 생각해 내고, 물의 양은 같게 해야 한다는 점을 먼저 말함.",
        3: "실험 설계 과정에서 대조군이 왜 필요한지 친구에게 쉬운 말로 설명함.",
        5: "가설을 쓰는 데 시간이 걸렸지만, 결과 예측은 표를 보고 차분하게 해석하려고 함.",
        7: "조건을 너무 많이 바꾸면 결과를 해석하기 어렵다는 점을 스스로 지적함.",
        9: "자료 해석에서는 부족한 부분이 있었지만, 실험 준비물 목록을 꼼꼼히 확인함.",
        11: "광합성량을 직접 볼 수 없다는 점을 질문하며 측정 방법의 필요성을 이해하려고 함.",
        13: "변인 통제 표를 다시 그려 보면서 다르게 해야 할 조건과 같게 해야 할 조건을 구분함.",
        15: "예상 결과가 틀릴까 봐 조심스러워했지만, 근거를 붙여 말하려는 태도가 좋았음.",
        17: "친구의 실험 설계를 듣고 빠진 통제 조건을 찾아 주는 모습이 있었음.",
        19: "광합성과 빛의 관계를 일상에서 본 식물 기르기 경험과 연결해 말함.",
        21: "그래프를 해석할 때 축 이름을 먼저 확인하는 습관이 생겨 자료 읽기가 안정됨.",
        23: "처음에는 결과를 감으로 예측했지만, 나중에는 표의 수치를 근거로 다시 고쳐 말함.",
        25: "실험 조건을 정리할 때 글보다 간단한 그림을 활용해 자기 방식으로 이해하려고 함.",
        27: "모둠 토의에서 나온 여러 조건을 칠판에 정리하며 활동 흐름을 잡는 데 도움을 줌.",
        29: "자료 해석은 아직 천천히 하는 편이나, 왜 그렇게 예측했는지 설명하려는 시도가 좋았음.",
    }

    def make_rows_for_area(area_id, levels, comments=None):
        comments = comments or {}
        rows = []
        pattern = level_patterns.get(area_id, levels)
        for idx, (class_no, number, name) in enumerate(students):
            level = pattern[idx % len(pattern)] if pattern else ""
            rows.append(
                {
                    "반": class_no,
                    "번호": number,
                    "성명": name,
                    "성취수준": level,
                    "추가 코멘트": comments.get(idx, ""),
                    "생성 수": 1,
                }
            )
        return rows

    student_rows_by_area = {}
    for area in areas:
        area_id = area["area_id"]
        if area_id == "mid_area_mix_comment":
            student_rows_by_area[area_id] = make_rows_for_area(area_id, ["총평"], mix_comments)
        elif area_id == "mid_area_photo_comment":
            student_rows_by_area[area_id] = make_rows_for_area(area_id, ["총평"], photo_comments)
        else:
            student_rows_by_area[area_id] = make_rows_for_area(area_id, area["levels"])

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
        "student_rows": student_rows_by_area[areas[0]["area_id"]],
        "student_rows_by_area": student_rows_by_area,
        "results": [],
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "app": "개꿀 생기부 - 중학교 간편 생성기",
        "version": "sample-mid-project-v17",
    }


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

    # JSON 배열로 온 경우 우선 처리
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

    # 한 줄에 여러 문장이 붙은 경우를 대비한 보조 분리
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

    base_clauses = []
    base_clauses.append(f"{name}에서 {rubric}")
    if unit:
        base_clauses.append(f"{unit} 단원의 {name} 활동에서 {rubric}")
    if desc:
        short_desc = re.sub(r"[.。]$", "", desc)
        base_clauses.append(f"{short_desc} 과정에서 {rubric}")
    base_clauses.append(f"{name} 활동을 통해 {rubric}")
    base_clauses.append(f"{name}의 수행 과정에서 {rubric}")
    base_clauses.append(f"{name}과 관련하여 {rubric}")
    base_clauses.append(f"{name} 활동에서 핵심 내용을 바탕으로 {rubric}")
    base_clauses.append(f"{name}을 수행하며 {rubric}")

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
        fallback = fallback_variations(area, level, count - len(sentences), extra_comment=extra_comment)
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
                "mode", "area_name", "level", "class", "number", "name", "final_text", "byte", "created_at", "generated_text", "extra_comment"
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
# UI 스타일
# =========================
st.markdown(
    """
    <style>
    div[data-testid="stExpander"] details:has(.mid-area-card) {
        background: linear-gradient(180deg, #EFF6FF 0%, #DBEAFE 100%) !important;
        border: 2px solid #93C5FD !important;
        border-radius: 18px !important;
        box-shadow: 0 5px 15px rgba(37, 99, 235, 0.10) !important;
        padding: 0.18rem 0.38rem 0.42rem 0.38rem !important;
        margin: 0.9rem 0 1.15rem 0 !important;
    }
    div[data-testid="stExpander"] details:has(.mid-area-card) > summary {
        background: #DBEAFE !important;
        border: 1px solid #BFDBFE !important;
        border-radius: 14px !important;
        margin: 0.2rem 0 0.55rem 0 !important;
        padding: 0.15rem 0.45rem !important;
    }
    div[data-testid="stExpander"] details:has(.mid-area-card) > summary p {
        color: #0F172A !important;
        font-weight: 900 !important;
    }
    .mid-area-card {display:none !important;}
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
# 사이드바
# =========================
with st.sidebar:
    st.header("중학교 간편 작업 관리")

    uploaded_mid_project = st.file_uploader(
        "중학교 프로젝트 JSON 불러오기",
        type=["json"],
        help="중학교 간편 생성기에서 저장한 JSON 파일을 다시 불러옵니다.",
    )
    if uploaded_mid_project and st.button("중학교 프로젝트 불러오기"):
        load_mid_project_json(uploaded_mid_project)
        st.success("중학교 프로젝트를 불러왔습니다.")
        st.rerun()

    st.download_button(
        "중학교 프로젝트 JSON 저장",
        data=mid_project_to_json(),
        file_name="개꿀생기부_mid_project.json",
        mime="application/json",
    )

    st.divider()

    if st.button("중학교 샘플 불러오기"):
        apply_mid_project_data(build_mid_sample_project())
        st.success("중학교 샘플 데이터를 불러왔습니다.")
        st.rerun()

    if st.button("중학교 결과만 초기화", type="secondary"):
        st.session_state.mid_results = []
        st.success("생성 결과를 초기화했습니다.")
        st.rerun()

    if st.button("중학교 페이지 전체 초기화", type="secondary"):
        for key in ["mid_settings", "mid_areas", "mid_results", "mid_student_rows", "mid_student_rows_by_area"]:
            if key in st.session_state:
                del st.session_state[key]
        init_mid_state()
        sanitize_mid_state()
        st.success("중학교 페이지를 초기화했습니다.")
        st.rerun()


# =========================
# 메인 화면
# =========================
st.title(MID_APP_TITLE)
st.caption(MID_APP_SUBTITLE)
st.info(
    "기존 상세형 페이지와 별도로 작동합니다. 중학교용은 활동·관찰 영역과 성취수준을 정한 뒤, "
    "같은 수준에 해당하는 문장을 여러 개 변주해서 만드는 구조입니다."
)

tab_basic, tab_area, tab_batch, tab_student, tab_results = st.tabs(
    ["① 기본 설정", "② 관찰 영역 설계", "③ 수준별 묶음 생성", "④ 학생별 배정 생성", "⑤ 결과 수정/다운로드"]
)


# =========================
# ① 기본 설정
# =========================
with tab_basic:
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
        settings["grade"] = st.text_input("학년", value=clean_text(settings.get("grade", "2")), key="mid_grade")
    with col4:
        settings["subject"] = st.text_input("과목명", value=clean_text(settings.get("subject", "과학")), key="mid_subject")

    col5, col6 = st.columns(2)
    with col5:
        settings["target_bytes_min"] = st.number_input(
            "목표 최소 byte",
            min_value=50,
            max_value=2000,
            value=int(settings.get("target_bytes_min", 250)),
            step=50,
            key="mid_target_min",
        )
    with col6:
        settings["target_bytes_max"] = st.number_input(
            "목표 최대 byte",
            min_value=50,
            max_value=2000,
            value=int(settings.get("target_bytes_max", 450)),
            step=50,
            key="mid_target_max",
        )

    st.markdown("#### 중학교용 작성 규칙")
    settings["custom_rules"] = st.text_area(
        "공통 작성 규칙",
        value=settings.get("custom_rules", MID_DEFAULT_RULES),
        height=220,
        key="mid_custom_rules",
    )

    st.success("기본 설정은 자동으로 세션에 반영됩니다. 왼쪽에서 JSON으로 저장할 수 있습니다.")


# =========================
# ② 관찰 영역 설계
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
    st.markdown("**성취수준 코드와 수준별 교사의 평가 문구**")
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


with tab_area:
    st.subheader("② 수행평가/관찰 영역 설계")
    st.caption("중학교용도 수행평가나 관찰 영역을 먼저 만들고, 그 안에 성취수준 코드와 평가 문구를 입력합니다.")

    with st.expander("➕ 새 수행평가/관찰 영역 추가", expanded=(len(st.session_state.mid_areas) == 0)):
        with st.form("mid_add_area_form"):
            col1, col2 = st.columns([2, 1])
            with col1:
                new_area_name = st.text_input("수행평가/관찰 영역명", placeholder="예: 소화 기관 모형 만들기")
                new_unit = st.text_input("단원/영역", placeholder="예: 동물의 몸과 영양소")
            with col2:
                new_use = st.checkbox("사용", value=True)
                st.caption("순서는 등록된 목록 순서대로 사용합니다.")
            new_desc = st.text_area(
                "활동 설명",
                placeholder="예: 소화 기관의 구조와 역할을 모형으로 표현하고 설명하는 활동",
                height=90,
            )
            submitted = st.form_submit_button("관찰 영역 추가")
            if submitted:
                if not clean_text(new_area_name):
                    st.warning("수행평가/관찰 영역명을 입력하세요.")
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
                    st.success("관찰 영역을 추가했습니다.")
                    st.rerun()

    if not st.session_state.mid_areas:
        st.info("아직 등록된 관찰 영역이 없습니다. 먼저 관찰 영역을 추가하거나 왼쪽의 샘플을 불러오세요.")
    else:
        for idx, area in enumerate(sort_areas(st.session_state.mid_areas), start=1):
            aid = area.get("area_id", "")
            levels = [clean_text(x) for x in area.get("levels", []) if clean_text(x)]
            with st.expander(f"📁 {area_label(area, idx)} · 성취수준 {len(levels)}개", expanded=True):
                st.markdown('<div class="mid-area-card"></div>', unsafe_allow_html=True)
                st.markdown(f"### 📁 {area_label(area, idx)}")

                col1, col2, col3 = st.columns([2, 2, 1])
                with col1:
                    area["name"] = st.text_input("관찰 영역명", value=area.get("name", ""), key=f"mid_area_name_{aid}")
                    area["unit"] = st.text_input("단원/영역", value=area.get("unit", ""), key=f"mid_area_unit_{aid}")
                with col2:
                    area["description"] = st.text_area("활동 설명", value=area.get("description", ""), height=112, key=f"mid_area_desc_{aid}")
                with col3:
                    area["use"] = st.checkbox("사용", value=area.get("use", True), key=f"mid_area_use_{aid}")
                    if st.button("삭제", key=f"mid_delete_area_{aid}"):
                        st.session_state.mid_areas = [x for x in st.session_state.mid_areas if x.get("area_id") != aid]
                        st.success("관찰 영역을 삭제했습니다.")
                        st.rerun()

                st.divider()
                new_levels, new_rubrics = render_level_input_block(
                    prefix=f"mid_area_levels_{aid}",
                    current_levels=area.get("levels", []),
                    current_rubrics=area.get("rubrics", {}),
                )
                if st.button("이 관찰 영역 성취수준 저장", type="primary", key=f"mid_save_levels_{aid}"):
                    area["levels"] = new_levels
                    area["rubrics"] = new_rubrics
                    sanitize_mid_state()
                    st.success("성취수준과 평가 문구를 저장했습니다.")
                    st.rerun()


# =========================
# AI 설정 공통 UI
# =========================
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
# ③ 수준별 묶음 생성
# =========================
with tab_batch:
    st.subheader("③ 수준별 문장 묶음 생성")
    st.caption("학생 이름 없이, 성취수준별로 여러 개의 문장을 한 번에 만듭니다. 나중에 필요한 학생에게 골라 붙이기 좋습니다.")

    if not st.session_state.mid_areas:
        st.warning("먼저 ②에서 수행평가/관찰 영역을 추가하세요.")
    else:
        selected_area = get_selected_area("mid_batch")
        if selected_area is None:
            st.warning("사용 설정된 관찰 영역이 없습니다.")
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
                            sentences, prompt = generate_variations(
                                selected_area,
                                level,
                                int(count),
                                provider,
                                api_key,
                                model,
                                variation_strength,
                            )
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
                    st.success(f"문장 {len(new_results)}개를 생성했습니다. ⑤ 결과 수정/다운로드에서 확인하세요.")
                    st.rerun()


# =========================
# ④ 학생별 배정 생성
# =========================
with tab_student:
    st.subheader("④ 학생별 배정 생성")
    st.caption("학생별로 성취수준과 간단한 추가 코멘트를 넣으면, 각 학생에게 붙일 수 있는 문장을 생성합니다. 학생 이름은 AI 입력 자료에 넣지 않습니다.")

    if not st.session_state.mid_areas:
        st.warning("먼저 ②에서 수행평가/관찰 영역을 추가하세요.")
    else:
        selected_area = get_selected_area("mid_student")
        if selected_area is None:
            st.warning("사용 설정된 관찰 영역이 없습니다.")
        else:
            levels = [clean_text(x) for x in selected_area.get("levels", []) if clean_text(x)]

            area_student_rows = get_student_rows_for_area(selected_area.get("area_id", ""), levels)

            edited_student_rows = st.data_editor(
                area_student_rows,
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
                    set_student_rows_for_area(selected_area.get("area_id", ""), edited_student_rows)
                    st.success("선택한 관찰 영역의 학생별 배정표를 저장했습니다.")
                    st.rerun()
            with col_hint:
                st.caption("성명은 결과표와 엑셀에만 표시됩니다. AI 프롬프트에는 이름, 반, 번호를 넣지 않습니다.")

            variation_strength = st.radio(
                "변주 강도",
                ["낮음", "보통", "높음"],
                index=1,
                horizontal=True,
                key="mid_student_variation_strength",
            )
            provider, model, api_key = render_ai_settings("mid_student_ai")

            if st.button("학생별 문장 생성", type="primary", use_container_width=True):
                set_student_rows_for_area(selected_area.get("area_id", ""), edited_student_rows)
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
                                selected_area,
                                level,
                                gen_count,
                                provider,
                                api_key,
                                model,
                                variation_strength,
                                extra_comment=extra_comment,
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
                    st.success(f"학생별 문장 {len(new_results)}개를 생성했습니다. ⑤ 결과 수정/다운로드에서 확인하세요.")
                    st.rerun()


# =========================
# ⑤ 결과 수정/다운로드
# =========================
with tab_results:
    st.subheader("⑤ 결과 수정 / 다운로드")

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

        edited_results = st.data_editor(
            results_df,
            use_container_width=True,
            height=520,
            hide_index=True,
            key="mid_results_editor",
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
                "final_text": st.column_config.TextColumn("최종 문장", width="large"),
                "byte": st.column_config.NumberColumn("byte", width="small"),
                "created_at": st.column_config.TextColumn("생성일시", width="medium"),
                "generated_text": st.column_config.TextColumn("생성 원문", width="large"),
                "extra_comment": st.column_config.TextColumn("추가 코멘트", width="large"),
                "prompt": None,
            },
        )

        col_save, col_clear = st.columns([1.4, 1.4])
        with col_save:
            if st.button("수정 문장 저장", type="primary", use_container_width=True):
                new_results = []
                for _, row in edited_results.iterrows():
                    row_dict = row.to_dict()
                    row_dict["final_text"] = clean_text(row_dict.get("final_text", ""))
                    row_dict["byte"] = byte_count(row_dict.get("final_text", ""))
                    new_results.append(row_dict)
                st.session_state.mid_results = new_results
                st.success("수정 문장을 저장했습니다.")
                st.rerun()
        with col_clear:
            if st.button("생성 결과 전체 삭제", type="secondary", use_container_width=True):
                st.session_state.mid_results = []
                st.warning("생성 결과를 모두 삭제했습니다.")
                st.rerun()

        st.divider()
        current_bytes = [byte_count(x.get("final_text", "")) for x in st.session_state.mid_results]
        col_metric1, col_metric2, col_metric3 = st.columns(3)
        with col_metric1:
            st.metric("생성 문장 수", len(st.session_state.mid_results))
        with col_metric2:
            st.metric("평균 byte", int(sum(current_bytes) / len(current_bytes)) if current_bytes else 0)
        with col_metric3:
            st.metric("최대 byte", max(current_bytes) if current_bytes else 0)

        excel_file = export_mid_excel()
        st.download_button(
            "📥 중학교 문장 결과 엑셀 다운로드",
            data=excel_file,
            file_name=f"개꿀생기부_mid_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        with st.expander("선택한 결과의 API 입력 프롬프트 확인", expanded=False):
            prompt_options = []
            prompt_map = {}
            for idx, result in enumerate(st.session_state.mid_results, start=1):
                label = f"{idx}. {result.get('mode', '')} / {result.get('area_name', '')} / {result.get('level', '')} / {clean_text(result.get('final_text', ''))[:32]}"
                prompt_options.append(label)
                prompt_map[label] = result.get("prompt", "")
            if prompt_options:
                selected_prompt_label = st.selectbox("프롬프트 선택", prompt_options, key="mid_prompt_view_select")
                st.text_area("API 입력 프롬프트", value=prompt_map.get(selected_prompt_label, ""), height=360)
