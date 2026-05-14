# ga4_media_auto_aggregator.py
# 사용법 1) 웹앱: streamlit run ga4_media_auto_aggregator.py
# 사용법 2) CLI:  python ga4_media_auto_aggregator.py 총사용자파일.csv 세션수파일.csv 결과.xlsx

import re
import sys
from io import BytesIO
from pathlib import Path

import pandas as pd


# 1) 공통 제외 기준: 세션 소스/매체에 아래 키워드가 포함된 행 제외
EXCLUDE_SOURCE_MEDIUM_KEYWORDS = [
    "brandsearch",
    "powercontents",
    "newproduct",
]

REQUIRED_BASE_COLS = ["세션 기본 채널 그룹", "세션 소스/매체", "이벤트 이름"]

ORGANIC_ORDER = ["Google", "Naver", "Daum", "Bing", "그외"]
REFERRAL_ORDER = ["KT·자사", "네이버", "커뮤니티·콘텐츠", "카카오", "그외"]
AI_SEARCH_ORDER = ["ChatGPT", "Gemini", "Perplexity"]


# 2) Referral 매체 분류 기준
# 필요하면 여기만 계속 보강하면 됨.
KT_OWNED_KEYWORDS = [
    "ktmmobile.com",
    "ktmmobile",
    "ktm모바일",
    "ktmyr.com",
    "ktmmarket.co.kr",
    "kt-aicc.com",
    "groupmail.kt.co.kr",
    "directmall",
    "kt.com",
    "kt.co.kr",
]

COMMUNITY_CONTENT_KEYWORDS = [
    # 커뮤니티
    "ppomppu",
    "dcinside",
    "fmkorea",
    "clien",
    "theqoo",
    "quasarzone",
    "reddit",
    "cetizen",

    # 콘텐츠/정보성 사이트
    "tistory",          # 중요: tistory는 카카오가 아니라 커뮤니티·콘텐츠
    "namu.wiki",
    "moyoplan",
    "mvnohub",
    "smartchoice",
    "dobiho",
    "funissu",
    "forloankr",
    "weayo",
    "rainygenius",
    "lunara",
    "yesteryear",
    "blog",             # naver/kakao 판정 이후 남은 일반 블로그성 도메인
]


def read_csv_safely(file):
    """GA4 CSV 인코딩이 섞여도 읽도록 처리."""
    encodings = ["utf-8-sig", "utf-8", "cp949", "euc-kr"]
    last_error = None

    for enc in encodings:
        try:
            return pd.read_csv(file, encoding=enc)
        except Exception as e:
            last_error = e

    raise ValueError(f"CSV 파일을 읽지 못했습니다: {last_error}")


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼명/문자값 앞뒤 공백 제거."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for col in REQUIRED_BASE_COLS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def validate_input(total_user_df: pd.DataFrame, metric_df: pd.DataFrame):
    missing_total = [c for c in REQUIRED_BASE_COLS + ["총 사용자"] if c not in total_user_df.columns]
    missing_metric = [c for c in REQUIRED_BASE_COLS + ["세션수"] if c not in metric_df.columns]

    if missing_total:
        raise ValueError(f"총 사용자 파일에 필요한 컬럼이 없습니다: {missing_total}")
    if missing_metric:
        raise ValueError(f"세션수 파일에 필요한 컬럼이 없습니다: {missing_metric}")


def auto_detect_files(df1: pd.DataFrame, df2: pd.DataFrame):
    """두 파일 중 총 사용자 파일/세션수 파일 자동 구분."""
    df1_cols = set(df1.columns)
    df2_cols = set(df2.columns)

    if "총 사용자" in df1_cols and "세션수" in df2_cols:
        return df1, df2
    if "총 사용자" in df2_cols and "세션수" in df1_cols:
        return df2, df1

    raise ValueError("두 파일 중 하나에는 '총 사용자', 다른 하나에는 '세션수' 컬럼이 있어야 합니다.")


def exact_not_excluded_mask(df: pd.DataFrame, keywords: list[str]) -> pd.Series:
    """세션 소스/매체에 제외 키워드가 포함된 행 제거. keywords가 비어있으면 전체 허용."""
    if not keywords:
        return pd.Series(True, index=df.index)
    source_medium = df["세션 소스/매체"].fillna("").astype(str).str.lower()
    pattern = "|".join(re.escape(k) for k in keywords)
    return ~source_medium.str.contains(pattern, regex=True)


def organic_channel_mask(df: pd.DataFrame) -> pd.Series:
    return df["세션 기본 채널 그룹"].str.lower().eq("organic search")


def referral_channel_mask(df: pd.DataFrame) -> pd.Series:
    allowed = {"referral", "organic social", "unassigned"}
    return df["세션 기본 채널 그룹"].str.lower().isin(allowed)


def ai_search_channel_mask(df: pd.DataFrame) -> pd.Series:
    return (
        df["세션 소스/매체"]
        .fillna("").astype(str).str.lower()
        .str.contains(r"gemini|gpt|perplexity", regex=True)
    )


def classify_ai_search(source_medium: str) -> str:
    s = str(source_medium).lower()
    if "gemini" in s:
        return "Gemini"
    if "gpt" in s:
        return "ChatGPT"
    if "perplexity" in s:
        return "Perplexity"
    return "그외"


def classify_organic(source_medium: str) -> str:
    s = str(source_medium).lower()

    if "google" in s:
        return "Google"
    if "naver" in s:
        return "Naver"
    if "daum" in s:
        return "Daum"
    if "bing" in s:
        return "Bing"
    return "그외"


def contains_any(text: str, keywords: list[str]) -> bool:
    return any(k in text for k in keywords)


def classify_referral(source_medium: str) -> str:
    s = str(source_medium).lower().strip()

    # tistory는 kakao 계열로 보지 않고 커뮤니티·콘텐츠로 우선 분류
    if "tistory" in s:
        return "커뮤니티·콘텐츠"

    # KT·자사
    if contains_any(s, KT_OWNED_KEYWORDS):
        return "KT·자사"

    # 네이버
    if "naver" in s:
        return "네이버"

    # 카카오: kakaochannel / social + daum.net 계열 + kakao.com 계열
    if (
        s == "kakaochannel / social"
        or "kakaochannel" in s
        or "kakao.com" in s
        or ".kakao.com" in s
        or "daum.net" in s
        or ".daum.net" in s
    ):
        return "카카오"

    # 커뮤니티·콘텐츠
    if contains_any(s, COMMUNITY_CONTENT_KEYWORDS):
        return "커뮤니티·콘텐츠"

    return "그외"


def aggregate_ga4(
    total_user_df: pd.DataFrame,
    metric_df: pd.DataFrame,
    channel_mask_func,
    classify_func,
    order: list[str],
    exclude_keywords: list[str] = EXCLUDE_SOURCE_MEDIUM_KEYWORDS,
) -> pd.DataFrame:
    """
    총사용자: 총 사용자 파일에서 이벤트 이름 = session_start 행만 사용
    세션수: 세션수 파일에서 이벤트 이름 = session_start
    작성_시작/작성_완료: 세션수 파일에서 이벤트 이름에 해당 문자열 포함
    """
    total_user_df = normalize_df(total_user_df)
    metric_df = normalize_df(metric_df)
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

    user_base["구분"] = user_base["세션 소스/매체"].apply(classify_func)
    metric_base["구분"] = metric_base["세션 소스/매체"].apply(classify_func)

    total_users = (
        user_base.groupby("구분", dropna=False)["총 사용자"]
        .sum()
        .rename("총사용자")
    )

    sessions = (
        metric_base[metric_base["이벤트 이름"].str.lower().eq("session_start")]
        .groupby("구분", dropna=False)["세션수"]
        .sum()
        .rename("세션수=session_start")
    )

    start_events = (
        metric_base[metric_base["이벤트 이름"].str.contains("작성_시작", na=False)]
        .groupby("구분", dropna=False)["세션수"]
        .sum()
        .rename("작성_시작 포함 이벤트수")
    )

    complete_events = (
        metric_base[metric_base["이벤트 이름"].str.contains("작성_완료", na=False)]
        .groupby("구분", dropna=False)["세션수"]
        .sum()
        .rename("작성_완료 포함 이벤트수")
    )

    result = pd.concat([total_users, sessions, start_events, complete_events], axis=1)
    result = result.reindex(order).fillna(0).astype(int)

    result.loc["합계"] = result.sum(numeric_only=True)
    result = result.reset_index().rename(columns={"index": "구분"})

    return result


def make_result(
    total_user_df: pd.DataFrame,
    metric_df: pd.DataFrame,
    exclude_keywords: list[str] = EXCLUDE_SOURCE_MEDIUM_KEYWORDS,
) -> dict[str, pd.DataFrame]:
    total_user_df, metric_df = auto_detect_files(
        normalize_df(total_user_df),
        normalize_df(metric_df),
    )

    kwargs = dict(
        total_user_df=total_user_df,
        metric_df=metric_df,
        exclude_keywords=exclude_keywords,
    )

    organic_result = aggregate_ga4(
        **kwargs,
        channel_mask_func=organic_channel_mask,
        classify_func=classify_organic,
        order=ORGANIC_ORDER,
    )

    referral_result = aggregate_ga4(
        **kwargs,
        channel_mask_func=referral_channel_mask,
        classify_func=classify_referral,
        order=REFERRAL_ORDER,
    )

    ai_search_result = aggregate_ga4(
        **kwargs,
        channel_mask_func=ai_search_channel_mask,
        classify_func=classify_ai_search,
        order=AI_SEARCH_ORDER,
    )

    return {
        "Organic Search": organic_result,
        "Referral_OS_Unassigned": referral_result,
        "AI Search": ai_search_result,
    }


def to_excel_bytes(results: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            safe_sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    return output.getvalue()


def run_cli():
    if len(sys.argv) < 4:
        print("사용법: python ga4_media_auto_aggregator.py 총사용자파일.csv 세션수파일.csv 결과.xlsx")
        sys.exit(1)

    file1 = sys.argv[1]
    file2 = sys.argv[2]
    output_path = sys.argv[3]

    df1 = read_csv_safely(file1)
    df2 = read_csv_safely(file2)

    results = make_result(df1, df2)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    print(f"완료: {output_path}")
    for name, df in results.items():
        print(f"\n[{name}]")
        print(df.to_string(index=False))


def copy_8rows_button(organic_df: pd.DataFrame, referral_df: pd.DataFrame):
    """Organic 합계 4개 + Referral 합계 4개를 세로 8행으로 복사."""
    import streamlit.components.v1 as components

    cols = ["총사용자", "세션수=session_start", "작성_시작 포함 이벤트수", "작성_완료 포함 이벤트수"]

    def get_total_row(df):
        total = df[df["구분"] == "합계"]
        if total.empty:
            return [0, 0, 0, 0]
        return [int(total.iloc[0][c]) for c in cols]

    organic_vals = get_total_row(organic_df)
    referral_vals = get_total_row(referral_df)

    eight_rows = "\n".join(str(v) for v in organic_vals + referral_vals)
    escaped = eight_rows.replace("`", "\\`")

    label_lines = [
        "Organic Search 총사용자", "Organic Search 세션수",
        "Organic Search 작성시작", "Organic Search 작성완료",
        "Referral 총사용자", "Referral 세션수",
        "Referral 작성시작", "Referral 작성완료",
    ]
    preview = "  /  ".join(
        f"{l}: {v}" for l, v in zip(label_lines, organic_vals + referral_vals)
    )

    components.html(
        f"""
        <div style="font-size:12px; color:#555; margin-bottom:4px;">{preview}</div>
        <button onclick="
            navigator.clipboard.writeText(`{escaped}`).then(() => {{
                this.textContent = '✅ 복사됨';
                setTimeout(() => this.textContent = '📋 8행 세로 복사 (Organic+Referral 합계)', 2000);
            }});
        " style="
            padding: 6px 16px;
            font-size: 13px;
            cursor: pointer;
            border: 1px solid #4a90d9;
            border-radius: 6px;
            background: #e8f0fe;
            color: #1a56a0;
            font-weight: bold;
        ">📋 8행 세로 복사 (Organic+Referral 합계)</button>
        """,
        height=60,
    )


def copy_button(df: pd.DataFrame, key: str):
    import streamlit.components.v1 as components

    tsv = df.to_csv(index=False, sep="\t")
    escaped = tsv.replace("`", "\\`")
    components.html(
        f"""
        <button onclick="
            navigator.clipboard.writeText(`{escaped}`).then(() => {{
                this.textContent = '✅ 복사됨';
                setTimeout(() => this.textContent = '📋 클립보드 복사 (엑셀 붙여넣기용)', 2000);
            }});
        " style="
            padding: 6px 14px;
            font-size: 13px;
            cursor: pointer;
            border: 1px solid #ccc;
            border-radius: 6px;
            background: #f8f9fa;
        ">📋 클립보드 복사 (엑셀 붙여넣기용)</button>
        """,
        height=40,
    )


def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="GA4 매체 자동 집계", layout="wide")
    st.title("GA4 매체 자동 집계")

    st.caption(
        "CSV 2개를 업로드하면 Organic Search와 Referral/Organic Social/Unassigned 기준을 자동 집계합니다."
    )

    with st.expander("세션 소스/매체 제외 필터 설정", expanded=False):
        use_filter = st.checkbox("제외 필터 사용", value=True)
        raw_input = st.text_area(
            "제외 키워드 (한 줄에 하나씩 입력)",
            value="\n".join(EXCLUDE_SOURCE_MEDIUM_KEYWORDS),
            height=120,
            disabled=not use_filter,
            help="세션 소스/매체에 해당 키워드가 포함된 행을 제외합니다.",
        )
        if use_filter:
            exclude_keywords = [k.strip() for k in raw_input.splitlines() if k.strip()]
        else:
            exclude_keywords = []

    uploaded_files = st.file_uploader(
        "총 사용자 CSV + 세션수 CSV를 날짜별로 여러 쌍 업로드하세요. (2개 단위, 짝수개)",
        type=["csv"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        return

    if len(uploaded_files) % 2 != 0:
        st.warning(f"파일은 짝수개(2, 4, 6…)로 업로드해야 합니다. 현재: {len(uploaded_files)}개")
        return

    try:
        dfs = [read_csv_safely(f) for f in uploaded_files]

        # 각 파일을 총사용자/세션수로 분류 후 각각 합산
        user_dfs, metric_dfs = [], []
        for df in dfs:
            cols = set(df.columns)
            if "총 사용자" in cols:
                user_dfs.append(df)
            elif "세션수" in cols:
                metric_dfs.append(df)
            else:
                raise ValueError(f"'총 사용자' 또는 '세션수' 컬럼이 없는 파일이 포함되어 있습니다.")

        if not user_dfs:
            raise ValueError("'총 사용자' 컬럼이 있는 파일이 없습니다.")
        if not metric_dfs:
            raise ValueError("'세션수' 컬럼이 있는 파일이 없습니다.")

        combined_user_df = pd.concat(user_dfs, ignore_index=True)
        combined_metric_df = pd.concat(metric_dfs, ignore_index=True)

        pair_info = f"총사용자 파일 {len(user_dfs)}개 · 세션수 파일 {len(metric_dfs)}개 합산"
        st.info(f"📂 {pair_info}")

        results = make_result(combined_user_df, combined_metric_df, exclude_keywords=exclude_keywords)

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

        st.subheader("AI Search (Gemini / GPT / Perplexity)")
        st.caption("세션 소스/매체에 gemini·gpt·perplexity 포함된 행 기준")
        st.dataframe(results["AI Search"], use_container_width=True)
        copy_button(results["AI Search"], key="ai")

        excel_bytes = to_excel_bytes(results)
        st.download_button(
            label="결과 Excel 다운로드",
            data=excel_bytes,
            file_name="ga4_media_aggregate_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(str(e))


if __name__ == "__main__":
    # streamlit run으로 실행하면 streamlit 화면, CLI 인자가 3개 있으면 CLI 실행
    if len(sys.argv) >= 4 and sys.argv[0].endswith(".py"):
        run_cli()
    else:
        run_streamlit()
