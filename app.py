# ga4_media_auto_aggregator.py
# 사용법 1) 웹앱: streamlit run app.py
# 사용법 2) CLI:  python app.py 파일A.csv 파일B.csv 결과.xlsx

import json
import re
import sys
from io import BytesIO

import pandas as pd


# ── 상수 ─────────────────────────────────────────────────────────────────────
EXCLUDE_SOURCE_MEDIUM_KEYWORDS = ["brandsearch", "newproduct"]

BASE_COLS = ["세션 기본 채널 그룹", "세션 소스/매체"]
EVENT_COL = "이벤트 이름"

# 집계 대상 이벤트 (정확히 일치)
START_EVENTS = [
    "esim_가입신청서_작성_시작__pc_공통_",
    "가입신청서_작성_시작__pc_공통_",
    "가입신청서_작성_시작__mo_공통_",
    "esim_가입신청서_작성_시작__mo_공통_",
]
COMPLETE_EVENTS = [
    "가입신청서_작성_완료__mo_공통_",
    "가입신청서_작성_완료__pc_공통_",
    "esim_가입신청서_작성_완료__mo_공통_",
    "esim_가입신청서_작성_완료__pc_공통_",
]
TARGET_EVENTS = START_EVENTS + COMPLETE_EVENTS

# ── 1차 구분 분류 상수 ─────────────────────────────────────────────────────
AI_GEO_SOURCES = [
    "chatgpt", "gemini", "perplexity", "copilot", "claude.ai", "manus.im", "doubao",
]
SOCIAL_MEDIA_SOURCES = [
    "facebook", "instagram", "l.instagram", "m.facebook", "l.facebook", "lm.facebook",
    "t.co", "tiktok", "threads", "l.threads", "zalo", "pinterest",
]
NAVER_SEO_EXACT_SOURCES = ["naver.com", "m.naver.com"]
NAVER_SEO_PARTIAL_SOURCES = [
    "blog.naver", "m.blog.naver", "cafe.naver", "m.cafe.naver",
    "kin.naver", "m.kin.naver", "blog.naverblogwidget", "m.search.naver.com",
]
NAVER_SEO_MEDIUMS = ["officialcafe", "blog_sp", "noticepost"]

PAYMENT_AUTH_SOURCES = [
    "mpay", "cert.", "checkplus", "orders.pay.naver", "nid.naver",
    "ekyc.naver", "login.microsoft", "shinhancard",
]

# 테이블 행 정렬 순서
ORGANIC_ORDER   = ["Google", "Naver", "Daum", "Bing", "그외"]
SEO_REF_ORDER   = ["AI/GEO", "Naver검색·컨텐츠", "커뮤니티·콘텐츠", "카카오", "그외"]
BIYEONG_ORDER   = ["KT·자사", "PPL·광고", "CRM", "결제·인증", "보류·출처불명", "기타"]

# ── 이전버전 (1차 구분 없음) 순서 ────────────────────────────────────────────
REFERRAL_ORDER_LEGACY  = ["KT·자사", "네이버", "커뮤니티·콘텐츠", "카카오", "그외"]
AI_SEARCH_ORDER_LEGACY = ["ChatGPT", "Gemini", "Perplexity"]

# 출력 컬럼명
COL_USERS    = "총사용자"
COL_SESSIONS = "세션수"
COL_START    = "작성_시작 이벤트수"
COL_COMPLETE = "작성_완료 이벤트수"

KT_OWNED_KEYWORDS = [
    "ktmmobile.com", "ktmmobile", "ktm모바일", "ktmyr.com",
    "ktmmarket.co.kr", "kt-aicc.com", "groupmail.kt.co.kr",
    "directmall", "kt.com", "kt.co.kr",
]
COMMUNITY_CONTENT_KEYWORDS = [
    "ppomppu", "dcinside", "fmkorea", "clien", "theqoo",
    "quasarzone", "reddit", "cetizen",
    "tistory", "namu.wiki", "moyoplan", "mvnohub", "smartchoice",
    "dobiho", "funissu", "forloankr", "weayo", "rainygenius",
    "lunara", "yesteryear", "blog",
]


# ── CSV 읽기 ──────────────────────────────────────────────────────────────────
def read_csv_safely(file):
    for enc in ["utf-8-sig", "utf-8", "cp949", "euc-kr"]:
        try:
            return pd.read_csv(file, encoding=enc)
        except Exception:
            pass
    raise ValueError("CSV 파일을 읽지 못했습니다.")


# ── 전처리 ────────────────────────────────────────────────────────────────────
def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for col in BASE_COLS + [EVENT_COL]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df


# ── 파일 형식 자동 감지 ───────────────────────────────────────────────────────
def detect_data_format(dfs: list[pd.DataFrame]):
    """
    신형: combined (세션수+총사용자) + event (주요이벤트)
    구형: user (총사용자+이벤트이름) + session (세션수+이벤트이름)
    """
    combined_list, event_list, user_list, session_list = [], [], [], []

    for df in dfs:
        df = normalize_df(df)
        cols = set(df.columns)
        has_users    = "총 사용자" in cols
        has_sessions = "세션수" in cols
        has_key_ev   = "주요 이벤트" in cols
        has_event    = EVENT_COL in cols

        if has_users and has_sessions:
            combined_list.append(df)
        elif has_key_ev and has_event:
            event_list.append(df)
        elif has_users and has_event:
            user_list.append(df)
        elif has_sessions and has_event:
            session_list.append(df)
        else:
            raise ValueError(
                f"인식할 수 없는 파일입니다. 컬럼: {sorted(cols)}\n"
                "필요: ① '세션수'+'총 사용자' 파일 + '주요 이벤트' 파일  "
                "또는  ② '총 사용자' 파일 + '세션수' 파일 (각각 이벤트 이름 포함)"
            )

    if combined_list and event_list:
        return "new", pd.concat(combined_list, ignore_index=True), pd.concat(event_list, ignore_index=True)
    if user_list and session_list:
        return "old", pd.concat(user_list, ignore_index=True), pd.concat(session_list, ignore_index=True)

    raise ValueError(
        "파일 조합을 인식할 수 없습니다.\n"
        "① 신형: '세션수+총사용자' 파일 + '주요이벤트' 파일\n"
        "② 구형: '총사용자(+이벤트이름)' 파일 + '세션수(+이벤트이름)' 파일"
    )


# ── 마스크 / 분류 함수 ─────────────────────────────────────────────────────────
def exact_not_excluded_mask(df, keywords):
    if not keywords:
        return pd.Series(True, index=df.index)
    src = df["세션 소스/매체"].fillna("").astype(str).str.lower()
    return ~src.str.contains("|".join(re.escape(k) for k in keywords), regex=True)

def organic_channel_mask(df):
    return df["세션 기본 채널 그룹"].str.lower().eq("organic search")

def _powercontents_mask(df):
    return df["세션 소스/매체"].fillna("").astype(str).str.lower().str.contains("powercontents", regex=False)

def organic_seo_mask(df):
    """Organic Search 채널 중 powercontents 제외 (SEO/GEO 연관용)"""
    return organic_channel_mask(df) & ~_powercontents_mask(df)

def organic_powercontents_mask(df):
    """Organic Search 채널 중 powercontents만 (비연관용)"""
    return organic_channel_mask(df) & _powercontents_mask(df)

def referral_channel_mask(df):
    return df["세션 기본 채널 그룹"].str.lower().isin({"referral", "organic social", "unassigned"})

def classify_organic(s):
    s = str(s).lower()
    if "google" in s: return "Google"
    if "naver"  in s: return "Naver"
    if "daum"   in s: return "Daum"
    if "bing"   in s: return "Bing"
    return "그외"


def _parse_src_med(s: str):
    s_lower = str(s).lower().strip()
    parts = s_lower.split(" / ")
    src = parts[0].strip()
    med = parts[1].strip() if len(parts) > 1 else ""
    return s_lower, src, med


def classify_referral_seo(s) -> str | None:
    """
    SEO/GEO 연관 referral 행의 3차구분 반환.
    비연관이면 None 반환.
    """
    s_lower, src, med = _parse_src_med(s)

    # ── 먼저 비연관 패턴 차단 (community 키워드 오탐 방지) ──
    if any(k in s_lower for k in KT_OWNED_KEYWORDS):
        return None
    if any(k in med for k in ["_ppl"]) or med in ["online_ad", "partner", "powercontents"]:
        return None
    if src == "crm":
        return None
    if any(x in src for x in PAYMENT_AUTH_SOURCES):
        return None
    if src == "(not set)" or "localhost" in src or (src and src[0].isdigit()):
        return None

    # ── SEO/GEO 판단 ──
    # AI/GEO
    if any(x in src for x in AI_GEO_SOURCES):
        return "AI/GEO"
    # 카카오 전체 → 별도 구분
    if "kakao" in src:
        return "카카오"
    # 소셜 미디어 → 커뮤니티·콘텐츠
    if any(x in src for x in SOCIAL_MEDIA_SOURCES) or (src == "ig" and med == "social") or med in ["social", "instagram_sp", "page"]:
        return "커뮤니티·콘텐츠"
    # Naver 검색/컨텐츠 (blog_ppl 제외)
    if "blog_ppl" not in med:
        if "m.search.naver.com" in src or (src == "naver" and med in NAVER_SEO_MEDIUMS):
            return "Naver검색·컨텐츠"
        if any(x in src for x in NAVER_SEO_PARTIAL_SOURCES):
            return "Naver검색·컨텐츠"
        if src in NAVER_SEO_EXACT_SOURCES:
            return "Naver검색·컨텐츠"
        if src == "naver_blog" and med in ["blog_sp", "noticepost"]:
            return "Naver검색·컨텐츠"
    # 커뮤니티/정보성
    if any(k in s_lower for k in COMMUNITY_CONTENT_KEYWORDS):
        return "커뮤니티·콘텐츠"

    # 위 비연관 패턴에 해당하지 않는 일반 referral → SEO/GEO 연관 그외
    return "그외"


def classify_referral_seo_safe(s) -> str:
    return classify_referral_seo(s) or "그외"


def classify_referral_biyeong(s) -> str:
    """비연관·보류 referral 행의 3차구분 반환."""
    s_lower, src, med = _parse_src_med(s)

    if any(k in s_lower for k in KT_OWNED_KEYWORDS):
        return "KT·자사"
    if any(k in med for k in ["_ppl"]) or med in ["online_ad", "partner", "powercontents"]:
        return "PPL·광고"
    if src == "crm":
        return "CRM"
    if any(x in src for x in PAYMENT_AUTH_SOURCES):
        return "결제·인증"
    if src == "(not set)" or "localhost" in src or (src and src[0].isdigit()):
        return "보류·출처불명"
    return "기타"


def ai_search_channel_mask(df):
    return df["세션 소스/매체"].fillna("").astype(str).str.lower().str.contains(
        r"gemini|gpt|perplexity", regex=True)

def classify_referral_legacy(s):
    s = str(s).lower().strip()
    if "tistory" in s: return "커뮤니티·콘텐츠"
    if any(k in s for k in KT_OWNED_KEYWORDS): return "KT·자사"
    if "naver" in s: return "네이버"
    if any(x in s for x in ("kakao", "daum.net", ".daum.net")): return "카카오"
    if any(k in s for k in COMMUNITY_CONTENT_KEYWORDS): return "커뮤니티·콘텐츠"
    return "그외"

def classify_ai_search(s):
    s = str(s).lower()
    if "gemini"     in s: return "Gemini"
    if "gpt"        in s: return "ChatGPT"
    if "perplexity" in s: return "Perplexity"
    return "그외"


def referral_seo_mask(df):
    base = referral_channel_mask(df)
    is_seo = df["세션 소스/매체"].apply(lambda s: classify_referral_seo(s) is not None)
    return base & is_seo


def referral_biyeong_mask(df):
    base = referral_channel_mask(df)
    is_biyeong = df["세션 소스/매체"].apply(lambda s: classify_referral_seo(s) is None)
    return base & is_biyeong



# ── 집계 ──────────────────────────────────────────────────────────────────────
def aggregate_ga4(fmt, df_a, df_b, channel_mask_func, classify_func, order,
                  exclude_keywords=EXCLUDE_SOURCE_MEDIUM_KEYWORDS, detail=False):
    """
    fmt=="new": df_a=combined(세션수+총사용자), df_b=event(주요이벤트)
    fmt=="old": df_a=total_user_df, df_b=metric_df (각각 이벤트이름 포함)
    detail=True 이면 [구분, 세션소스/매체] 단위로 집계
    """
    df_a = normalize_df(df_a)
    df_b = normalize_df(df_b)

    grp_cols = ["구분", "세션 소스/매체"] if detail else ["구분"]

    if fmt == "new":
        combined = df_a[channel_mask_func(df_a) & exact_not_excluded_mask(df_a, exclude_keywords)].copy()
        events   = df_b[channel_mask_func(df_b) & exact_not_excluded_mask(df_b, exclude_keywords)].copy()

        combined["구분"] = combined["세션 소스/매체"].apply(classify_func)
        events["구분"]   = events["세션 소스/매체"].apply(classify_func)

        total_users = combined.groupby(grp_cols, dropna=False)["총 사용자"].sum().rename(COL_USERS)
        sessions    = combined.groupby(grp_cols, dropna=False)["세션수"].sum().rename(COL_SESSIONS)
        start_ev    = (
            events[events[EVENT_COL].isin(START_EVENTS)]
            .groupby(grp_cols, dropna=False)["주요 이벤트"].sum().rename(COL_START)
        )
        complete_ev = (
            events[events[EVENT_COL].isin(COMPLETE_EVENTS)]
            .groupby(grp_cols, dropna=False)["주요 이벤트"].sum().rename(COL_COMPLETE)
        )

    else:  # old
        user_base = df_a[
            channel_mask_func(df_a)
            & exact_not_excluded_mask(df_a, exclude_keywords)
            & df_a[EVENT_COL].str.lower().eq("session_start")
        ].copy()
        metric_base = df_b[
            channel_mask_func(df_b)
            & exact_not_excluded_mask(df_b, exclude_keywords)
        ].copy()

        user_base["구분"]   = user_base["세션 소스/매체"].apply(classify_func)
        metric_base["구분"] = metric_base["세션 소스/매체"].apply(classify_func)

        total_users = user_base.groupby(grp_cols, dropna=False)["총 사용자"].sum().rename(COL_USERS)
        sessions    = (
            metric_base[metric_base[EVENT_COL].str.lower().eq("session_start")]
            .groupby(grp_cols, dropna=False)["세션수"].sum().rename(COL_SESSIONS)
        )
        start_ev    = (
            metric_base[metric_base[EVENT_COL].isin(START_EVENTS)]
            .groupby(grp_cols, dropna=False)["세션수"].sum().rename(COL_START)
        )
        complete_ev = (
            metric_base[metric_base[EVENT_COL].isin(COMPLETE_EVENTS)]
            .groupby(grp_cols, dropna=False)["세션수"].sum().rename(COL_COMPLETE)
        )

    result = pd.concat([total_users, sessions, start_ev, complete_ev], axis=1)

    if detail:
        result = result.fillna(0).astype(int).reset_index()
        return result.sort_values(["구분", COL_USERS], ascending=[True, False]).reset_index(drop=True)
    else:
        result = result.reindex(order).fillna(0).astype(int)
        result.loc["합계"] = result.sum(numeric_only=True)
        return result.reset_index().rename(columns={"index": "구분"})


def make_result(dfs: list[pd.DataFrame], exclude_keywords=EXCLUDE_SOURCE_MEDIUM_KEYWORDS):
    fmt, df_a, df_b = detect_data_format(dfs)
    kw = dict(fmt=fmt, df_a=df_a, df_b=df_b, exclude_keywords=exclude_keywords)

    def agg(mask, cls, order, detail=False):
        return aggregate_ga4(**kw, channel_mask_func=mask, classify_func=cls, order=order, detail=detail)

    return {
        "Organic Search":            agg(organic_seo_mask,          classify_organic,            ORGANIC_ORDER),
        "Organic Search_detail":     agg(organic_seo_mask,          classify_organic,            ORGANIC_ORDER,  detail=True),
        "SEO Referral":              agg(referral_seo_mask,         classify_referral_seo_safe,  SEO_REF_ORDER),
        "SEO Referral_detail":       agg(referral_seo_mask,         classify_referral_seo_safe,  SEO_REF_ORDER,  detail=True),
        "AI Search":                 agg(ai_search_channel_mask,    classify_ai_search,          AI_SEARCH_ORDER_LEGACY),
        "AI Search_detail":          agg(ai_search_channel_mask,    classify_ai_search,          AI_SEARCH_ORDER_LEGACY, detail=True),
        "Organic 비연관":             agg(organic_powercontents_mask, classify_referral_biyeong,  BIYEONG_ORDER),
        "Organic 비연관_detail":      agg(organic_powercontents_mask, classify_referral_biyeong,  BIYEONG_ORDER,  detail=True),
        "비연관·보류":               agg(referral_biyeong_mask,      classify_referral_biyeong,   BIYEONG_ORDER),
        "비연관·보류_detail":         agg(referral_biyeong_mask,      classify_referral_biyeong,   BIYEONG_ORDER,  detail=True),
    }


def make_result_legacy(dfs: list[pd.DataFrame], exclude_keywords=EXCLUDE_SOURCE_MEDIUM_KEYWORDS):
    """1차 구분(SEO/GEO) 없이 기존 3개 테이블 구조로 집계."""
    fmt, df_a, df_b = detect_data_format(dfs)
    kw = dict(fmt=fmt, df_a=df_a, df_b=df_b, exclude_keywords=exclude_keywords)

    def agg(mask, cls, order, detail=False):
        return aggregate_ga4(**kw, channel_mask_func=mask, classify_func=cls, order=order, detail=detail)

    return {
        "Organic Search":                agg(organic_channel_mask,   classify_organic,        ORGANIC_ORDER),
        "Organic Search_detail":         agg(organic_channel_mask,   classify_organic,        ORGANIC_ORDER,  detail=True),
        "Referral_OS_Unassigned":        agg(referral_channel_mask,  classify_referral_legacy, REFERRAL_ORDER_LEGACY),
        "Referral_OS_Unassigned_detail": agg(referral_channel_mask,  classify_referral_legacy, REFERRAL_ORDER_LEGACY, detail=True),
        "AI Search":                     agg(ai_search_channel_mask, classify_ai_search,      AI_SEARCH_ORDER_LEGACY),
        "AI Search_detail":              agg(ai_search_channel_mask, classify_ai_search,      AI_SEARCH_ORDER_LEGACY, detail=True),
    }


def copy_all_button_legacy(results: dict):
    import streamlit.components.v1 as components
    sections = [
        ("Organic Search",                    add_cvr(results["Organic Search"])),
        ("Referral / Organic Social / Unassigned", add_cvr(results["Referral_OS_Unassigned"])),
        ("AI Search (ChatGPT / Gemini / Perplexity)", add_cvr(results["AI Search"])),
    ]
    combined = "\n\n".join(
        f"[ {lbl} ]\n{df[df['구분'] != '합계'].to_csv(index=False, sep=chr(9), header=False)}"
        for lbl, df in sections
    )
    combined = combined.replace("`", "'")
    import streamlit.components.v1 as components
    components.html(
        f"""<button onclick="
            navigator.clipboard.writeText(`{combined}`).then(() => {{
                this.textContent = '✅ 전체 복사됨';
                setTimeout(() => this.textContent = '📋 전체 표 복사 (3개 한번에)', 2000);
            }});" style="padding:8px 20px;font-size:14px;cursor:pointer;
            border:1px solid #2e7d32;border-radius:6px;background:#e8f5e9;
            color:#1b5e20;font-weight:bold;width:100%;">
            📋 전체 표 복사 (3개 한번에)</button>""",
        height=48,
    )


def show_results_legacy(results: dict, st, key_prefix: str = ""):
    copy_all_button_legacy(results)
    st.divider()

    os_df  = add_cvr(results["Organic Search"])
    ref_df = add_cvr(results["Referral_OS_Unassigned"])
    ai_df  = add_cvr(results["AI Search"])

    st.subheader("Organic Search")
    st.dataframe(os_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(os_df, key=f"{key_prefix}l_organic")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["Organic Search_detail"]),
            file_name="legacy_organic_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_l_organic_detail")

    st.subheader("Referral / Organic Social / Unassigned")
    st.dataframe(ref_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(ref_df, key=f"{key_prefix}l_referral")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["Referral_OS_Unassigned_detail"]),
            file_name="legacy_referral_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_l_referral_detail")

    st.divider()
    st.markdown("**Organic + Referral 합계 8행 복사**")
    copy_8rows_button(results["Organic Search"], results["Referral_OS_Unassigned"])

    st.divider()
    st.subheader("AI Search (ChatGPT / Gemini / Perplexity)")
    st.caption("세션 소스/매체에 gpt·gemini·perplexity 포함된 행 기준")
    st.dataframe(ai_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(ai_df, key=f"{key_prefix}l_ai")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["AI Search_detail"]),
            file_name="legacy_ai_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_l_ai_detail")

    st.download_button("결과 Excel 다운로드",
        data=to_excel_bytes(results),
        file_name="ga4_legacy_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key_prefix}dl_l_excel")


# ── GA4 API (OAuth) ───────────────────────────────────────────────────────────
def fetch_ga4_data_oauth(credentials, property_id: str, start_date: str, end_date: str):
    from google.analytics.data_v1beta import BetaAnalyticsDataClient
    from google.analytics.data_v1beta.types import (
        DateRange, Dimension, Filter, FilterExpression,
        FilterExpressionList, Metric, RunReportRequest,
    )

    client     = BetaAnalyticsDataClient(credentials=credentials)
    prop       = f"properties/{property_id}"
    date_range = DateRange(start_date=start_date, end_date=end_date)

    base_dims = [
        Dimension(name="date"),
        Dimension(name="sessionDefaultChannelGroup"),
        Dimension(name="sessionSourceMedium"),
        Dimension(name="deviceCategory"),
    ]

    combined_req = RunReportRequest(
        property=prop, date_ranges=[date_range],
        dimensions=base_dims,
        metrics=[Metric(name="sessions"), Metric(name="totalUsers")],
        limit=100000,
    )

    event_req = RunReportRequest(
        property=prop, date_ranges=[date_range],
        dimensions=base_dims + [Dimension(name="eventName")],
        metrics=[Metric(name="keyEvents")],
        dimension_filter=FilterExpression(filter=Filter(
            field_name="eventName",
            in_list_filter=Filter.InListFilter(
                values=TARGET_EVENTS,
                case_sensitive=True,
            ),
        )),
        limit=100000,
    )

    def to_df(response, metric_cols):
        dim_names = [h.name for h in response.dimension_headers]
        rows = [
            [d.value for d in row.dimension_values] + [m.value for m in row.metric_values]
            for row in response.rows
        ]
        col_map = {
            "date": "date",
            "sessionDefaultChannelGroup": "세션 기본 채널 그룹",
            "sessionSourceMedium": "세션 소스/매체",
            "deviceCategory": "기기 카테고리",
            "eventName": "이벤트 이름",
        }
        cols = [col_map.get(n, n) for n in dim_names] + metric_cols
        df   = pd.DataFrame(rows, columns=cols)
        for c in metric_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        return df

    combined_df = to_df(client.run_report(combined_req), ["세션수", "총 사용자"])
    event_df    = to_df(client.run_report(event_req),    ["주요 이벤트"])
    return combined_df, event_df


# ── 엑셀 ──────────────────────────────────────────────────────────────────────
def to_excel_bytes(results):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in results.items():
            if not name.endswith("_detail"):
                df.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()

def detail_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="상세", index=False)
    return output.getvalue()


# ── CVR 컬럼 추가 ─────────────────────────────────────────────────────────────
def add_cvr(df: pd.DataFrame) -> pd.DataFrame:
    """합계 포함 전 행에 CVR1·CVR2 컬럼 추가 후 컬럼 순서 재정렬."""
    df = df.copy()

    def pct(num, den):
        return f"{num / den * 100:.1f}%" if den > 0 else "-"

    df["CVR1"] = df.apply(lambda r: pct(r[COL_START],    r[COL_SESSIONS]), axis=1)
    df["CVR2"] = df.apply(lambda r: pct(r[COL_COMPLETE], r[COL_START]),    axis=1)

    cols = ["구분", COL_USERS, COL_SESSIONS, COL_START, "CVR1", COL_COMPLETE, "CVR2"]
    return df[cols]


# ── Streamlit UI 헬퍼 ─────────────────────────────────────────────────────────
def copy_button(df: pd.DataFrame, key: str):
    import streamlit.components.v1 as components
    data_only = df[df["구분"] != "합계"]
    tsv = data_only.to_csv(index=False, sep="\t", header=False).replace("`", "'")
    components.html(
        f"""<button onclick="
            navigator.clipboard.writeText(`{tsv}`).then(() => {{
                this.textContent = '✅ 복사됨';
                setTimeout(() => this.textContent = '📋 클립보드 복사 (엑셀 붙여넣기용)', 2000);
            }});" style="padding:6px 14px;font-size:13px;cursor:pointer;
            border:1px solid #ccc;border-radius:6px;background:#f8f9fa;">
            📋 클립보드 복사 (엑셀 붙여넣기용)</button>""",
        height=40,
    )


def copy_8rows_button(organic_df: pd.DataFrame, seo_ref_df: pd.DataFrame):
    import streamlit.components.v1 as components
    cols = [COL_USERS, COL_SESSIONS, COL_START, COL_COMPLETE]

    def total_vals(df):
        row = df[df["구분"] == "합계"]
        return [int(row.iloc[0][c]) if not row.empty else 0 for c in cols]

    vals   = total_vals(organic_df) + total_vals(seo_ref_df)
    text   = "\n".join(str(v) for v in vals).replace("`", "'")
    labels = [
        "Organic 총사용자", "Organic 세션수", "Organic 작성시작", "Organic 작성완료",
        "SEO Referral 총사용자", "SEO Referral 세션수", "SEO Referral 작성시작", "SEO Referral 작성완료",
    ]
    preview = "  /  ".join(f"{l}: {v}" for l, v in zip(labels, vals))
    components.html(
        f"""<div style="font-size:12px;color:#555;margin-bottom:4px;">{preview}</div>
        <button onclick="
            navigator.clipboard.writeText(`{text}`).then(() => {{
                this.textContent = '✅ 복사됨';
                setTimeout(() => this.textContent = '📋 8행 세로 복사 (Organic + SEO Referral 합계)', 2000);
            }});" style="padding:6px 16px;font-size:13px;cursor:pointer;
            border:1px solid #4a90d9;border-radius:6px;background:#e8f0fe;
            color:#1a56a0;font-weight:bold;">
            📋 8행 세로 복사 (Organic + SEO Referral 합계)</button>""",
        height=60,
    )


def copy_all_button(results: dict):
    import streamlit.components.v1 as components
    sections = [
        ("SEO/GEO 연관 > Organic Search",                    add_cvr(results["Organic Search"])),
        ("SEO/GEO 연관 > Referral (AI/GEO·Naver·커뮤니티·카카오)", add_cvr(results["SEO Referral"])),
        ("SEO/GEO 연관 > AI Search",                         add_cvr(results["AI Search"])),
        ("비연관·보류 > Organic Search 비연관",               add_cvr(results["Organic 비연관"])),
        ("비연관·보류 > Referral",                           add_cvr(results["비연관·보류"])),
    ]
    combined = "\n\n".join(
        f"[ {lbl} ]\n{df[df['구분'] != '합계'].to_csv(index=False, sep=chr(9), header=False)}"
        for lbl, df in sections
    )
    combined = combined.replace("`", "'")
    components.html(
        f"""<button onclick="
            navigator.clipboard.writeText(`{combined}`).then(() => {{
                this.textContent = '✅ 전체 복사됨';
                setTimeout(() => this.textContent = '📋 전체 표 복사 (5개 한번에)', 2000);
            }});" style="padding:8px 20px;font-size:14px;cursor:pointer;
            border:1px solid #2e7d32;border-radius:6px;background:#e8f5e9;
            color:#1b5e20;font-weight:bold;width:100%;">
            📋 전체 표 복사 (5개 한번에)</button>""",
        height=48,
    )


def show_results(results: dict, st, key_prefix: str = ""):
    copy_all_button(results)
    st.divider()

    # ── SEO/GEO 연관 유입 ──────────────────────────────────────────────────
    st.markdown("### 🟢 SEO/GEO 연관 유입")

    os_df        = add_cvr(results["Organic Search"])
    seo_df       = add_cvr(results["SEO Referral"])
    os_bi_df     = add_cvr(results["Organic 비연관"])
    biyeong_df   = add_cvr(results["비연관·보류"])

    st.subheader("Organic Search")
    st.dataframe(os_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(os_df, key=f"{key_prefix}organic")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["Organic Search_detail"]),
            file_name="organic_search_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_organic_detail")

    st.subheader("Referral (AI/GEO · Naver검색/컨텐츠 · 커뮤니티·콘텐츠 · 카카오)")
    st.caption("Referral 채널 중 SEO/GEO 연관 유입 (AI 검색, 네이버 검색/블로그/카페, 커뮤니티, 소셜, 카카오)")
    st.dataframe(seo_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(seo_df, key=f"{key_prefix}seo_ref")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["SEO Referral_detail"]),
            file_name="seo_referral_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_seo_ref_detail")

    st.divider()
    st.markdown("**Organic + SEO Referral 합계 8행 복사**")
    copy_8rows_button(results["Organic Search"], results["SEO Referral"])

    # ── AI Search ──────────────────────────────────────────────────────────
    if "AI Search" in results:
        st.divider()
        st.markdown("### 🤖 AI Search (ChatGPT / Gemini / Perplexity)")
        st.caption("세션 소스/매체에 gpt·gemini·perplexity 포함된 행 기준")
        ai_df = add_cvr(results["AI Search"])
        st.dataframe(ai_df, use_container_width=True)
        col1, col2 = st.columns([2, 1])
        with col1:
            copy_button(ai_df, key=f"{key_prefix}ai_search")
        with col2:
            st.download_button("📥 상세 데이터 다운로드",
                data=detail_excel_bytes(results["AI Search_detail"]),
                file_name="ai_search_detail.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{key_prefix}dl_ai_search_detail")

    # ── 비연관·보류 ────────────────────────────────────────────────────────
    st.divider()
    st.markdown("### 🔴 비연관·보류")

    st.subheader("Organic Search 비연관")
    st.caption("Organic Search 채널 중 비연관 유입 (naver/powercontents 등)")
    st.dataframe(os_bi_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(os_bi_df, key=f"{key_prefix}os_biyeong")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["Organic 비연관_detail"]),
            file_name="organic_biyeong_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_os_biyeong_detail")

    st.subheader("Referral 비연관 (KT·자사 · PPL·광고 · CRM · 결제·인증 · 보류)")
    st.caption("Referral 채널 중 SEO/GEO 비연관 유입 (자사, 광고운영, CRM, 결제/인증, 출처불명)")
    st.dataframe(biyeong_df, use_container_width=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        copy_button(biyeong_df, key=f"{key_prefix}biyeong")
    with col2:
        st.download_button("📥 상세 데이터 다운로드",
            data=detail_excel_bytes(results["비연관·보류_detail"]),
            file_name="biyeongwan_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}dl_biyeong_detail")

    st.divider()
    st.download_button("결과 Excel 다운로드",
        data=to_excel_bytes(results),
        file_name="ga4_media_aggregate_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key_prefix}dl_excel")


# ── Streamlit 메인 ────────────────────────────────────────────────────────────
def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="GA4 매체 자동 집계", layout="wide")
    st.title("GA4 매체 자동 집계")

    # ── 제외 필터 (공통) ──────────────────────────────────────────────────────
    with st.expander("세션 소스/매체 제외 필터 설정", expanded=False):
        use_filter = st.checkbox("제외 필터 사용", value=True)
        raw_kw = st.text_area(
            "제외 키워드 (한 줄에 하나씩)",
            value="\n".join(EXCLUDE_SOURCE_MEDIUM_KEYWORDS),
            height=90, disabled=not use_filter,
            help="세션 소스/매체에 해당 키워드가 포함된 행을 제외합니다.",
        )
        exclude_keywords = (
            [k.strip() for k in raw_kw.splitlines() if k.strip()] if use_filter else []
        )

    tab_file, tab_api, tab_legacy = st.tabs(["📁 파일 업로드", "🔗 GA4 API 연동", "📂 이전버전 (GA4 API)"])

    # ── 탭 1: 파일 업로드 ─────────────────────────────────────────────────────
    with tab_file:
        st.caption(
            "**신형** (권장): ① 세션수+총사용자 파일  +  ② 주요이벤트 파일\n\n"
            "**구형**: ① 총사용자(이벤트이름 포함) 파일  +  ② 세션수(이벤트이름 포함) 파일\n\n"
            "날짜별 여러 쌍 업로드 시 같은 타입끼리 자동 합산합니다."
        )
        uploaded = st.file_uploader(
            "CSV 파일 업로드 (2개 이상)",
            type=["csv"], accept_multiple_files=True, key="file_uploader",
        )

        if not uploaded:
            st.info("CSV 파일을 업로드하면 집계 결과가 표시됩니다.")
        elif len(uploaded) < 2:
            st.warning("최소 2개 파일을 업로드해야 합니다.")
        else:
            try:
                dfs = [read_csv_safely(f) for f in uploaded]
                results = make_result(dfs, exclude_keywords=exclude_keywords)
                st.info(f"📂 파일 {len(uploaded)}개 처리 완료")
                show_results(results, st, key_prefix="file_")
            except Exception as e:
                st.error(str(e))

    # ── 탭 2: GA4 API 연동 ───────────────────────────────────────────────────
    with tab_api:
        import datetime
        from streamlit_oauth import OAuth2Component
        from google.oauth2.credentials import Credentials as OAuthCredentials

        secret_prop_id = st.secrets.get("GA4_PROPERTY_ID", "")
        client_id      = st.secrets.get("GOOGLE_CLIENT_ID", "")
        client_secret  = st.secrets.get("GOOGLE_CLIENT_SECRET", "")

        if not client_id or not client_secret:
            st.error("Secrets에 GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET 가 설정되지 않았습니다.")
            st.stop()

        col_l, col_r = st.columns([1, 1])
        with col_l:
            property_id = st.text_input(
                "GA4 Property ID", value=str(secret_prop_id),
                placeholder="예: 123456789", help="GA4 관리 → 속성 → 속성 ID (숫자)",
            )
        with col_r:
            today      = datetime.date.today()
            start_date = st.date_input("시작일", value=today.replace(day=1))
            end_date   = st.date_input("종료일", value=today)

        oauth2 = OAuth2Component(
            client_id=client_id, client_secret=client_secret,
            authorize_endpoint="https://accounts.google.com/o/oauth2/v2/auth",
            token_endpoint="https://oauth2.googleapis.com/token",
            refresh_token_endpoint="https://oauth2.googleapis.com/token",
            revoke_token_endpoint="https://oauth2.googleapis.com/revoke",
        )

        access_token  = (st.session_state.get("oauth_token") or {}).get("access_token")
        refresh_token = (st.session_state.get("oauth_token") or {}).get("refresh_token")

        if not access_token:
            try:
                token_result = oauth2.authorize_button(
                    name="🔑 Google 계정으로 로그인",
                    redirect_uri=st.secrets.get("REDIRECT_URI", "http://localhost:8501"),
                    scope="https://www.googleapis.com/auth/analytics.readonly",
                    extras_params={"access_type": "offline", "prompt": "consent"},
                    key="google_oauth", use_container_width=False,
                )
                if token_result and "token" in token_result:
                    st.session_state["oauth_token"] = token_result["token"]
                    st.rerun()
            except Exception:
                stale = [k for k in st.session_state if any(
                    x in k.lower() for x in ("state", "code", "google_oauth"))]
                for k in stale:
                    del st.session_state[k]
                st.warning("로그인 세션이 만료되었습니다. 다시 로그인해주세요.")
                st.rerun()

        access_token  = (st.session_state.get("oauth_token") or {}).get("access_token")
        refresh_token = (st.session_state.get("oauth_token") or {}).get("refresh_token")

        if access_token:
            col_s, col_lo = st.columns([4, 1])
            with col_s:
                st.success("✅ Google 로그인 완료")
            with col_lo:
                if st.button("로그아웃", key="logout_btn"):
                    del st.session_state["oauth_token"]
                    st.rerun()

            if st.button("📡 데이터 가져오기", type="primary", disabled=not property_id):
                st.session_state.pop("api_results", None)
                st.session_state.pop("api_results_info", None)
                if start_date > end_date:
                    st.error("시작일이 종료일보다 늦습니다.")
                else:
                    try:
                        with st.spinner("GA4 API 호출 중…"):
                            creds = OAuthCredentials(
                                token=access_token, refresh_token=refresh_token,
                                token_uri="https://oauth2.googleapis.com/token",
                                client_id=client_id, client_secret=client_secret,
                            )
                            combined_df, event_df = fetch_ga4_data_oauth(
                                credentials=creds,
                                property_id=str(property_id).strip(),
                                start_date=start_date.strftime("%Y-%m-%d"),
                                end_date=end_date.strftime("%Y-%m-%d"),
                            )
                        st.session_state["api_results"] = make_result(
                            [combined_df, event_df], exclude_keywords=exclude_keywords
                        )
                        st.session_state["api_results_info"] = (
                            f"combined {len(combined_df):,}행 · event {len(event_df):,}행 수신 완료"
                        )
                    except Exception as e:
                        err = str(e)
                        if "401" in err or "credentials" in err.lower():
                            st.warning("인증이 만료되었습니다. 다시 로그인해주세요.")
                            del st.session_state["oauth_token"]
                            st.rerun()
                        else:
                            st.error(err)

            if "api_results" in st.session_state:
                st.success(st.session_state.get("api_results_info", ""))
                show_results(st.session_state["api_results"], st, key_prefix="api_")
        else:
            st.info("위 버튼으로 Google 계정에 로그인하면 GA4 데이터를 가져올 수 있습니다.")

    # ── 탭 3: 이전버전 (GA4 API, 1차 구분 없음) ──────────────────────────────
    with tab_legacy:
        import datetime as _dt
        st.caption("연관/비연관 구분 없이 기존 3개 표 구조 (Organic Search / Referral / AI Search)")

        secret_prop_id2 = st.secrets.get("GA4_PROPERTY_ID", "")
        client_id2      = st.secrets.get("GOOGLE_CLIENT_ID", "")
        client_secret2  = st.secrets.get("GOOGLE_CLIENT_SECRET", "")

        if not client_id2 or not client_secret2:
            st.error("Secrets에 GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET 가 설정되지 않았습니다.")
        else:
            col_l2, col_r2 = st.columns([1, 1])
            with col_l2:
                property_id2 = st.text_input(
                    "GA4 Property ID", value=str(secret_prop_id2),
                    placeholder="예: 123456789",
                    key="legacy_property_id",
                )
            with col_r2:
                today2      = _dt.date.today()
                start_date2 = st.date_input("시작일", value=today2.replace(day=1), key="legacy_start")
                end_date2   = st.date_input("종료일", value=today2, key="legacy_end")

            # OAuth: GA4 API 연동 탭의 토큰 공유
            legacy_access  = (st.session_state.get("oauth_token") or {}).get("access_token")
            legacy_refresh = (st.session_state.get("oauth_token") or {}).get("refresh_token")

            if not legacy_access:
                st.warning("GA4 API 연동 탭에서 Google 계정으로 먼저 로그인해주세요.")
            else:
                st.success("✅ Google 로그인 완료 (GA4 API 연동 탭과 공유)")

                if st.button("📡 이전버전 데이터 가져오기", type="primary",
                             disabled=not property_id2, key="legacy_fetch_btn"):
                    if start_date2 > end_date2:
                        st.error("시작일이 종료일보다 늦습니다.")
                    else:
                        try:
                            from google.oauth2.credentials import Credentials as OAuthCredentials
                            with st.spinner("GA4 API 호출 중…"):
                                creds2 = OAuthCredentials(
                                    token=legacy_access, refresh_token=legacy_refresh,
                                    token_uri="https://oauth2.googleapis.com/token",
                                    client_id=client_id2, client_secret=client_secret2,
                                )
                                combined_df2, event_df2 = fetch_ga4_data_oauth(
                                    credentials=creds2,
                                    property_id=str(property_id2).strip(),
                                    start_date=start_date2.strftime("%Y-%m-%d"),
                                    end_date=end_date2.strftime("%Y-%m-%d"),
                                )
                            st.session_state["legacy_results"] = make_result_legacy(
                                [combined_df2, event_df2], exclude_keywords=exclude_keywords
                            )
                            st.session_state["legacy_results_info"] = (
                                f"combined {len(combined_df2):,}행 · event {len(event_df2):,}행 수신 완료"
                            )
                        except Exception as e:
                            err = str(e)
                            if "401" in err or "credentials" in err.lower():
                                st.warning("인증이 만료되었습니다. GA4 API 연동 탭에서 다시 로그인해주세요.")
                                del st.session_state["oauth_token"]
                                st.rerun()
                            else:
                                st.error(err)

                if "legacy_results" in st.session_state:
                    st.success(st.session_state.get("legacy_results_info", ""))
                    show_results_legacy(st.session_state["legacy_results"], st, key_prefix="leg_")


# ── CLI ───────────────────────────────────────────────────────────────────────
def run_cli():
    if len(sys.argv) < 4:
        print("사용법: python app.py 파일A.csv 파일B.csv 결과.xlsx")
        sys.exit(1)
    dfs = [read_csv_safely(f) for f in sys.argv[1:-1]]
    results = make_result(dfs)
    with pd.ExcelWriter(sys.argv[-1], engine="openpyxl") as writer:
        for name, df in results.items():
            if not name.endswith("_detail"):
                df.to_excel(writer, sheet_name=name[:31], index=False)
    print(f"완료: {sys.argv[-1]}")
    for name, df in results.items():
        if not name.endswith("_detail"):
            print(f"\n[{name}]\n{df.to_string(index=False)}")


if __name__ == "__main__":
    if len(sys.argv) >= 4 and sys.argv[0].endswith(".py"):
        run_cli()
    else:
        run_streamlit()
