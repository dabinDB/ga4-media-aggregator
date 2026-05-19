# ga4_media_auto_aggregator.py
# 사용법 1) 웹앱: streamlit run app.py
# 사용법 2) CLI:  python app.py 총사용자파일.csv 세션수파일.csv 결과.xlsx

import json
import re
import sys
from io import BytesIO

import pandas as pd


# ── 공통 제외 기준 ────────────────────────────────────────────────────────────
EXCLUDE_SOURCE_MEDIUM_KEYWORDS = [
    "brandsearch",
    "powercontents",
    "newproduct",
]

REQUIRED_BASE_COLS = ["세션 기본 채널 그룹", "세션 소스/매체", "이벤트 이름"]

ORGANIC_ORDER  = ["Google", "Naver", "Daum", "Bing", "그외"]
REFERRAL_ORDER = ["KT·자사", "네이버", "커뮤니티·콘텐츠", "카카오", "그외"]
AI_SEARCH_ORDER = ["ChatGPT", "Gemini", "Perplexity"]

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


# ── GA4 API 데이터 가져오기 ───────────────────────────────────────────────────
def fetch_ga4_data(
    credentials_info: dict,
    property_id: str,
    start_date: str,
    end_date: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    GA4 Data API로 두 보고서를 가져와 (총사용자 df, 세션수 df) 반환.
    - 총사용자: eventName CONTAINS 'session_start'
    - 세션수: eventName PARTIAL_REGEXP '가입신청서|session_start|유심_배송신청서'
              AND eventName PARTIAL_REGEXP '가입신청서|session_start'
    """
    from google.analytics.data_v1beta import BetaAnalyticsDataClient
    from google.analytics.data_v1beta.types import (
        DateRange, Dimension, Filter, FilterExpression,
        FilterExpressionList, Metric, RunReportRequest,
    )
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        credentials_info,
        scopes=["https://www.googleapis.com/auth/analytics.readonly"],
    )
    client = BetaAnalyticsDataClient(credentials=creds)

    prop = f"properties/{property_id}"
    date_range = DateRange(start_date=start_date, end_date=end_date)
    dimensions = [
        Dimension(name="date"),
        Dimension(name="sessionDefaultChannelGroup"),
        Dimension(name="sessionSourceMedium"),
        Dimension(name="eventName"),
    ]

    # ① 총사용자 보고서
    users_req = RunReportRequest(
        property=prop,
        date_ranges=[date_range],
        dimensions=dimensions,
        metrics=[Metric(name="totalUsers")],
        dimension_filter=FilterExpression(
            filter=Filter(
                field_name="eventName",
                string_filter=Filter.StringFilter(
                    match_type=Filter.StringFilter.MatchType.CONTAINS,
                    value="session_start",
                ),
            )
        ),
        limit=100000,
    )

    # ② 세션수 보고서
    sessions_req = RunReportRequest(
        property=prop,
        date_ranges=[date_range],
        dimensions=dimensions,
        metrics=[Metric(name="sessions")],
        dimension_filter=FilterExpression(
            and_group=FilterExpressionList(
                expressions=[
                    FilterExpression(
                        filter=Filter(
                            field_name="eventName",
                            string_filter=Filter.StringFilter(
                                match_type=Filter.StringFilter.MatchType.PARTIAL_REGEXP,
                                value="가입신청서|session_start|유심_배송신청서",
                            ),
                        )
                    ),
                    FilterExpression(
                        filter=Filter(
                            field_name="eventName",
                            string_filter=Filter.StringFilter(
                                match_type=Filter.StringFilter.MatchType.PARTIAL_REGEXP,
                                value="가입신청서|session_start",
                            ),
                        )
                    ),
                ]
            )
        ),
        limit=100000,
    )

    def response_to_df(response, metric_col: str) -> pd.DataFrame:
        rows = [
            [d.value for d in row.dimension_values] + [m.value for m in row.metric_values]
            for row in response.rows
        ]
        cols = ["date", "세션 기본 채널 그룹", "세션 소스/매체", "이벤트 이름", metric_col]
        df = pd.DataFrame(rows, columns=cols)
        df[metric_col] = pd.to_numeric(df[metric_col], errors="coerce").fillna(0).astype(int)
        return df

    users_df   = response_to_df(client.run_report(users_req),   "총 사용자")
    sessions_df = response_to_df(client.run_report(sessions_req), "세션수")
    return users_df, sessions_df


def fetch_ga4_data_oauth(credentials, property_id: str, start_date: str, end_date: str):
    """OAuth 액세스 토큰으로 GA4 Data API 호출 (서비스 계정 불필요)."""
    from google.analytics.data_v1beta import BetaAnalyticsDataClient
    from google.analytics.data_v1beta.types import (
        DateRange, Dimension, Filter, FilterExpression,
        FilterExpressionList, Metric, RunReportRequest,
    )

    client = BetaAnalyticsDataClient(credentials=credentials)

    prop       = f"properties/{property_id}"
    date_range = DateRange(start_date=start_date, end_date=end_date)
    dimensions = [
        Dimension(name="date"),
        Dimension(name="sessionDefaultChannelGroup"),
        Dimension(name="sessionSourceMedium"),
        Dimension(name="eventName"),
    ]

    users_req = RunReportRequest(
        property=prop, date_ranges=[date_range], dimensions=dimensions,
        metrics=[Metric(name="totalUsers")],
        dimension_filter=FilterExpression(filter=Filter(
            field_name="eventName",
            string_filter=Filter.StringFilter(
                match_type=Filter.StringFilter.MatchType.CONTAINS,
                value="session_start",
            ),
        )),
        limit=100000,
    )

    sessions_req = RunReportRequest(
        property=prop, date_ranges=[date_range], dimensions=dimensions,
        metrics=[Metric(name="sessions")],
        dimension_filter=FilterExpression(
            and_group=FilterExpressionList(expressions=[
                FilterExpression(filter=Filter(
                    field_name="eventName",
                    string_filter=Filter.StringFilter(
                        match_type=Filter.StringFilter.MatchType.PARTIAL_REGEXP,
                        value="가입신청서|session_start|유심_배송신청서",
                    ),
                )),
                FilterExpression(filter=Filter(
                    field_name="eventName",
                    string_filter=Filter.StringFilter(
                        match_type=Filter.StringFilter.MatchType.PARTIAL_REGEXP,
                        value="가입신청서|session_start",
                    ),
                )),
            ])
        ),
        limit=100000,
    )

    def to_df(response, metric_col):
        rows = [
            [d.value for d in row.dimension_values] + [m.value for m in row.metric_values]
            for row in response.rows
        ]
        df = pd.DataFrame(rows, columns=["date", "세션 기본 채널 그룹", "세션 소스/매체", "이벤트 이름", metric_col])
        df[metric_col] = pd.to_numeric(df[metric_col], errors="coerce").fillna(0).astype(int)
        return df

    return (
        to_df(client.run_report(users_req),    "총 사용자"),
        to_df(client.run_report(sessions_req), "세션수"),
    )


# ── 전처리 / 분류 ─────────────────────────────────────────────────────────────
def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for col in REQUIRED_BASE_COLS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df


def validate_input(total_user_df, metric_df):
    miss_u = [c for c in REQUIRED_BASE_COLS + ["총 사용자"] if c not in total_user_df.columns]
    miss_m = [c for c in REQUIRED_BASE_COLS + ["세션수"]    if c not in metric_df.columns]
    if miss_u:
        raise ValueError(f"총 사용자 파일에 필요한 컬럼이 없습니다: {miss_u}")
    if miss_m:
        raise ValueError(f"세션수 파일에 필요한 컬럼이 없습니다: {miss_m}")


def auto_detect_files(df1, df2):
    if "총 사용자" in df1.columns and "세션수" in df2.columns:
        return df1, df2
    if "총 사용자" in df2.columns and "세션수" in df1.columns:
        return df2, df1
    raise ValueError("하나에는 '총 사용자', 다른 하나에는 '세션수' 컬럼이 있어야 합니다.")


def exact_not_excluded_mask(df, keywords):
    if not keywords:
        return pd.Series(True, index=df.index)
    src = df["세션 소스/매체"].fillna("").astype(str).str.lower()
    pattern = "|".join(re.escape(k) for k in keywords)
    return ~src.str.contains(pattern, regex=True)


def organic_channel_mask(df):
    return df["세션 기본 채널 그룹"].str.lower().eq("organic search")

def referral_channel_mask(df):
    return df["세션 기본 채널 그룹"].str.lower().isin({"referral", "organic social", "unassigned"})

def ai_search_channel_mask(df):
    return df["세션 소스/매체"].fillna("").astype(str).str.lower().str.contains(
        r"gemini|gpt|perplexity", regex=True
    )


def classify_organic(s):
    s = str(s).lower()
    if "google" in s:    return "Google"
    if "naver" in s:     return "Naver"
    if "daum" in s:      return "Daum"
    if "bing" in s:      return "Bing"
    return "그외"

def classify_referral(s):
    s = str(s).lower().strip()
    if "tistory" in s:   return "커뮤니티·콘텐츠"
    if any(k in s for k in KT_OWNED_KEYWORDS): return "KT·자사"
    if "naver" in s:     return "네이버"
    if any(x in s for x in ("kakaochannel", "kakao.com", ".kakao.com", "daum.net", ".daum.net")):
        return "카카오"
    if any(k in s for k in COMMUNITY_CONTENT_KEYWORDS): return "커뮤니티·콘텐츠"
    return "그외"

def classify_ai_search(s):
    s = str(s).lower()
    if "gemini" in s:    return "Gemini"
    if "gpt" in s:       return "ChatGPT"
    if "perplexity" in s: return "Perplexity"
    return "그외"


# ── 집계 ──────────────────────────────────────────────────────────────────────
def aggregate_ga4(
    total_user_df, metric_df,
    channel_mask_func, classify_func, order,
    exclude_keywords=EXCLUDE_SOURCE_MEDIUM_KEYWORDS,
):
    total_user_df = normalize_df(total_user_df)
    metric_df     = normalize_df(metric_df)
    validate_input(total_user_df, metric_df)

    user_base = total_user_df[
        channel_mask_func(total_user_df)
        & exact_not_excluded_mask(total_user_df, exclude_keywords)
        & total_user_df["이벤트 이름"].str.lower().eq("session_start")
    ].copy()

    metric_base = metric_df[
        channel_mask_func(metric_df)
        & exact_not_excluded_mask(metric_df, exclude_keywords)
    ].copy()

    user_base["구분"]   = user_base["세션 소스/매체"].apply(classify_func)
    metric_base["구분"] = metric_base["세션 소스/매체"].apply(classify_func)

    total_users = (
        user_base.groupby("구분", dropna=False)["총 사용자"]
        .sum().rename("총사용자")
    )
    sessions = (
        metric_base[metric_base["이벤트 이름"].str.lower().eq("session_start")]
        .groupby("구분", dropna=False)["세션수"]
        .sum().rename("세션수=session_start")
    )
    start_events = (
        metric_base[metric_base["이벤트 이름"].str.contains("작성_시작", na=False)]
        .groupby("구분", dropna=False)["세션수"]
        .sum().rename("작성_시작 포함 이벤트수")
    )
    complete_events = (
        metric_base[metric_base["이벤트 이름"].str.contains("작성_완료", na=False)]
        .groupby("구분", dropna=False)["세션수"]
        .sum().rename("작성_완료 포함 이벤트수")
    )

    result = pd.concat([total_users, sessions, start_events, complete_events], axis=1)
    result = result.reindex(order).fillna(0).astype(int)
    result.loc["합계"] = result.sum(numeric_only=True)
    return result.reset_index().rename(columns={"index": "구분"})


def make_result(total_user_df, metric_df, exclude_keywords=EXCLUDE_SOURCE_MEDIUM_KEYWORDS):
    total_user_df, metric_df = auto_detect_files(
        normalize_df(total_user_df), normalize_df(metric_df)
    )
    kw = dict(total_user_df=total_user_df, metric_df=metric_df, exclude_keywords=exclude_keywords)
    return {
        "Organic Search":       aggregate_ga4(**kw, channel_mask_func=organic_channel_mask,  classify_func=classify_organic,    order=ORGANIC_ORDER),
        "Referral_OS_Unassigned": aggregate_ga4(**kw, channel_mask_func=referral_channel_mask, classify_func=classify_referral,   order=REFERRAL_ORDER),
        "AI Search":            aggregate_ga4(**kw, channel_mask_func=ai_search_channel_mask, classify_func=classify_ai_search,  order=AI_SEARCH_ORDER),
    }


def to_excel_bytes(results):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in results.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


# ── Streamlit UI 헬퍼 ─────────────────────────────────────────────────────────
def copy_button(df: pd.DataFrame, key: str):
    import streamlit.components.v1 as components
    tsv = df.to_csv(index=False, sep="\t").replace("`", "\\`")
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


def copy_8rows_button(organic_df: pd.DataFrame, referral_df: pd.DataFrame):
    import streamlit.components.v1 as components
    cols = ["총사용자", "세션수=session_start", "작성_시작 포함 이벤트수", "작성_완료 포함 이벤트수"]

    def total_vals(df):
        row = df[df["구분"] == "합계"]
        return [int(row.iloc[0][c]) if not row.empty else 0 for c in cols]

    vals = total_vals(organic_df) + total_vals(referral_df)
    text = "\n".join(str(v) for v in vals).replace("`", "\\`")
    labels = [
        "Organic 총사용자", "Organic 세션수", "Organic 작성시작", "Organic 작성완료",
        "Referral 총사용자", "Referral 세션수", "Referral 작성시작", "Referral 작성완료",
    ]
    preview = "  /  ".join(f"{l}: {v}" for l, v in zip(labels, vals))
    components.html(
        f"""<div style="font-size:12px;color:#555;margin-bottom:4px;">{preview}</div>
        <button onclick="
            navigator.clipboard.writeText(`{text}`).then(() => {{
                this.textContent = '✅ 복사됨';
                setTimeout(() => this.textContent = '📋 8행 세로 복사 (Organic+Referral 합계)', 2000);
            }});" style="padding:6px 16px;font-size:13px;cursor:pointer;
            border:1px solid #4a90d9;border-radius:6px;background:#e8f0fe;
            color:#1a56a0;font-weight:bold;">
            📋 8행 세로 복사 (Organic+Referral 합계)</button>""",
        height=60,
    )


def copy_all_button(results: dict):
    """3개 표 전체를 표 이름 포함해서 한 번에 복사."""
    import streamlit.components.v1 as components

    sections = [
        ("Organic Search",                results["Organic Search"]),
        ("Referral / Organic Social / Unassigned", results["Referral_OS_Unassigned"]),
        ("AI Search (ChatGPT / Gemini / Perplexity)", results["AI Search"]),
    ]

    blocks = []
    for label, df in sections:
        tsv = df.to_csv(index=False, sep="\t")
        blocks.append(f"[ {label} ]\n{tsv}")

    combined = "\n\n".join(blocks).replace("`", "\\`").replace("\\", "\\\\").replace("`", "\\`")
    # 역슬래시 이중 처리 없이 단순하게
    combined = "\n\n".join(blocks).replace("`", "'")

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


def show_results(results: dict, st):
    """집계 결과를 화면에 출력하는 공통 함수."""
    copy_all_button(results)
    st.divider()

    st.subheader("Organic Search")
    st.dataframe(results["Organic Search"], use_container_width=True)
    copy_button(results["Organic Search"], key="organic")

    st.subheader("Referral / Organic Social / Unassigned")
    st.dataframe(results["Referral_OS_Unassigned"], use_container_width=True)
    copy_button(results["Referral_OS_Unassigned"], key="referral")

    st.divider()
    st.markdown("**Organic + Referral 합계 8행 복사**")
    copy_8rows_button(results["Organic Search"], results["Referral_OS_Unassigned"])
    st.divider()

    st.subheader("AI Search (ChatGPT / Gemini / Perplexity)")
    st.caption("세션 소스/매체에 gpt·gemini·perplexity 포함된 행 기준")
    st.dataframe(results["AI Search"], use_container_width=True)
    copy_button(results["AI Search"], key="ai")

    excel_bytes = to_excel_bytes(results)
    st.download_button(
        label="결과 Excel 다운로드",
        data=excel_bytes,
        file_name="ga4_media_aggregate_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Streamlit 메인 ────────────────────────────────────────────────────────────
def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="GA4 매체 자동 집계", layout="wide")
    st.title("GA4 매체 자동 집계")

    # 제외 필터 (공통)
    with st.expander("세션 소스/매체 제외 필터 설정", expanded=False):
        use_filter = st.checkbox("제외 필터 사용", value=True)
        raw_kw = st.text_area(
            "제외 키워드 (한 줄에 하나씩)",
            value="\n".join(EXCLUDE_SOURCE_MEDIUM_KEYWORDS),
            height=110,
            disabled=not use_filter,
            help="세션 소스/매체에 해당 키워드가 포함된 행을 제외합니다.",
        )
        exclude_keywords = (
            [k.strip() for k in raw_kw.splitlines() if k.strip()]
            if use_filter else []
        )

    tab_file, tab_api = st.tabs(["📁 파일 업로드", "🔗 GA4 API 연동"])

    # ── 탭 1: 파일 업로드 ─────────────────────────────────────────────────────
    with tab_file:
        uploaded = st.file_uploader(
            "총 사용자 CSV + 세션수 CSV를 날짜별로 여러 쌍 업로드 (짝수개)",
            type=["csv"],
            accept_multiple_files=True,
            key="file_uploader",
        )

        if not uploaded:
            st.info("CSV 파일을 업로드하면 집계 결과가 표시됩니다.")
        elif len(uploaded) % 2 != 0:
            st.warning(f"파일은 짝수개로 업로드해야 합니다. 현재: {len(uploaded)}개")
        else:
            try:
                dfs = [read_csv_safely(f) for f in uploaded]
                user_dfs, metric_dfs = [], []
                for df in dfs:
                    if "총 사용자" in df.columns:
                        user_dfs.append(df)
                    elif "세션수" in df.columns:
                        metric_dfs.append(df)
                    else:
                        raise ValueError("'총 사용자' 또는 '세션수' 컬럼이 없는 파일이 있습니다.")

                if not user_dfs:
                    raise ValueError("'총 사용자' 파일이 없습니다.")
                if not metric_dfs:
                    raise ValueError("'세션수' 파일이 없습니다.")

                st.info(f"📂 총사용자 {len(user_dfs)}개 · 세션수 {len(metric_dfs)}개 합산")
                results = make_result(
                    pd.concat(user_dfs, ignore_index=True),
                    pd.concat(metric_dfs, ignore_index=True),
                    exclude_keywords=exclude_keywords,
                )
                show_results(results, st)
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

        # ── 날짜 & Property ID ───────────────────────────────────────────────
        col_l, col_r = st.columns([1, 1])
        with col_l:
            property_id = st.text_input(
                "GA4 Property ID",
                value=str(secret_prop_id),
                placeholder="예: 123456789",
                help="GA4 관리 → 속성 → 속성 ID (숫자)",
            )
        with col_r:
            today      = datetime.date.today()
            start_date = st.date_input("시작일", value=today.replace(day=1))
            end_date   = st.date_input("종료일", value=today)

        # ── Google OAuth 로그인 ──────────────────────────────────────────────
        oauth2 = OAuth2Component(
            client_id=client_id,
            client_secret=client_secret,
            authorize_endpoint="https://accounts.google.com/o/oauth2/v2/auth",
            token_endpoint="https://oauth2.googleapis.com/token",
            refresh_token_endpoint="https://oauth2.googleapis.com/token",
            revoke_token_endpoint="https://oauth2.googleapis.com/revoke",
        )

        # 이미 로그인된 경우 authorize_button 호출 자체를 건너뜀
        access_token = (st.session_state.get("oauth_token") or {}).get("access_token")

        if not access_token:
            try:
                token_result = oauth2.authorize_button(
                    name="🔑 Google 계정으로 로그인",
                    redirect_uri=st.secrets.get("REDIRECT_URI", "http://localhost:8501"),
                    scope="https://www.googleapis.com/auth/analytics.readonly",
                    key="google_oauth",
                    use_container_width=False,
                )
                if token_result and "token" in token_result:
                    st.session_state["oauth_token"] = token_result["token"]
                    st.rerun()
            except Exception:
                # state 불일치 등 OAuth flow 오류 → flow 관련 키만 초기화
                stale_keys = [k for k in st.session_state if any(
                    x in k.lower() for x in ("state", "code", "google_oauth")
                )]
                for k in stale_keys:
                    del st.session_state[k]
                st.warning("로그인 세션이 만료되었습니다. 다시 로그인해주세요.")
                st.rerun()

        access_token = (st.session_state.get("oauth_token") or {}).get("access_token")

        if access_token:
            col_status, col_logout = st.columns([4, 1])
            with col_status:
                st.success("✅ Google 로그인 완료")
            with col_logout:
                if st.button("로그아웃", key="logout_btn"):
                    del st.session_state["oauth_token"]
                    st.rerun()

            fetch_btn = st.button(
                "📡 데이터 가져오기", type="primary",
                disabled=not property_id,
            )

            if fetch_btn:
                if start_date > end_date:
                    st.error("시작일이 종료일보다 늦습니다.")
                else:
                    try:
                        with st.spinner("GA4 API 호출 중…"):
                            creds = OAuthCredentials(token=access_token)
                            users_df, sessions_df = fetch_ga4_data_oauth(
                                credentials=creds,
                                property_id=str(property_id).strip(),
                                start_date=start_date.strftime("%Y-%m-%d"),
                                end_date=end_date.strftime("%Y-%m-%d"),
                            )
                        st.success(
                            f"총사용자 {len(users_df):,}행 · 세션수 {len(sessions_df):,}행 수신 완료"
                        )
                        results = make_result(users_df, sessions_df, exclude_keywords=exclude_keywords)
                        show_results(results, st)
                    except Exception as e:
                        st.error(str(e))
        else:
            st.info("위 버튼으로 Google 계정에 로그인하면 GA4 데이터를 가져올 수 있습니다.")


# ── CLI ───────────────────────────────────────────────────────────────────────
def run_cli():
    if len(sys.argv) < 4:
        print("사용법: python app.py 총사용자.csv 세션수.csv 결과.xlsx")
        sys.exit(1)
    df1 = read_csv_safely(sys.argv[1])
    df2 = read_csv_safely(sys.argv[2])
    results = make_result(df1, df2)
    with pd.ExcelWriter(sys.argv[3], engine="openpyxl") as writer:
        for name, df in results.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    print(f"완료: {sys.argv[3]}")
    for name, df in results.items():
        print(f"\n[{name}]\n{df.to_string(index=False)}")


if __name__ == "__main__":
    if len(sys.argv) >= 4 and sys.argv[0].endswith(".py"):
        run_cli()
    else:
        run_streamlit()
