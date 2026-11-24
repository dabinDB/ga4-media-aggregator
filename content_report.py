# -*- coding: utf-8 -*-
import datetime as dt
import html
import re
import zipfile
from collections import defaultdict
from io import BytesIO

import pandas as pd

PROPERTY_ID = "310948970"
WEEK1_START = dt.date(2026, 11, 9)
N_WEEKS = 8

# 값은 엑셀 '콘텐츠' 열과 정확히 일치해야 함
CONTENTS = {
    "1884": "알뜰폰이란? 가입 전에 알아두면 좋은 것들",
    "1886": "알뜰폰 가입 전 무엇을 확인해야 하나요?",
    "1888": "알뜰폰 어떻게 시작하나요? 요금제 선택과 개통",
}
URL_REGEX = r"tips\.do\?(.*&)?ntcartSeq={seq}(&.*)?$"

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
CONVERSION_METRIC = "keyEvents"

AI_SOURCES = [
    ("ChatGPT", r"chatgpt|openai"),
    ("Perplexity", r"perplexity"),
    ("Gemini", r"gemini|bard"),
    ("기타 AI", r"claude|copilot|deepseek|grok|you\.com|wrtn|mistral|meta\.ai|poe\.com|phind|liner|clova-x|cue\.naver"),
]
AI_CHANNEL = "AI Assistant"

CHANNELS = ["Direct", "Organic Search", "Referral", "AI Assistant", "Organic Social"]
OTHER_CHANNEL = "기타"

SHEET_CONTENT = "콘텐츠별 성과"
SHEET_CHANNEL = "채널별 전환"

CONTENT_COLS = ["오가닉 세션", "ChatGPT", "Perplexity", "Gemini", "기타 AI", "세션", "참여 세션", "사용자",
                "활성 사용자", "페이지뷰 수", "평균 참여율", "평균 참여시간(초)", "유입 세션", "가입시작", "가입완료"]
CHANNEL_COLS = ["유입 세션", "가입시작", "가입완료"]

# 수식 열을 덮어쓰지 않도록 리포트 시트의 연속된 입력 칸 단위로 복사
CONTENT_COPY_BLOCKS = [
    ("G:K", ["오가닉 세션", "ChatGPT", "Perplexity", "Gemini", "기타 AI"]),
    ("M:Q", ["세션", "참여 세션", "사용자", "활성 사용자", "페이지뷰 수"]),
    ("S:V", ["평균 참여율", "평균 참여시간(초)", "유입 세션", "가입시작"]),
    ("X", ["가입완료"]),
]
CHANNEL_COPY_BLOCKS = [
    ("D:E", ["유입 세션", "가입시작"]),
    ("G", ["가입완료"]),
]


# ── GA4 ──────────────────────────────────────────────────────────────────────
def _filter(field, regex):
    from google.analytics.data_v1beta.types import Filter, FilterExpression
    sf = Filter.StringFilter(match_type=Filter.StringFilter.MatchType.PARTIAL_REGEXP, value=regex)
    return FilterExpression(filter=Filter(field_name=field, string_filter=sf))


def _in_list(field, values):
    from google.analytics.data_v1beta.types import Filter, FilterExpression
    return FilterExpression(filter=Filter(
        field_name=field,
        in_list_filter=Filter.InListFilter(values=values, case_sensitive=True),
    ))


def _and(*exprs):
    from google.analytics.data_v1beta.types import FilterExpression, FilterExpressionList
    return FilterExpression(and_group=FilterExpressionList(expressions=list(exprs)))


def run(client, property_id, start, end, dims, mets, flt):
    from google.analytics.data_v1beta.types import DateRange, Dimension, Metric, RunReportRequest
    rows, offset = [], 0
    while True:
        req = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[DateRange(start_date=start.isoformat(), end_date=end.isoformat())],
            dimensions=[Dimension(name=d) for d in dims],
            metrics=[Metric(name=m) for m in mets],
            dimension_filter=flt, limit=100000, offset=offset)
        res = client.run_report(req)
        for r in res.rows:
            row = {d: v.value for d, v in zip(dims, r.dimension_values)}
            row.update({m: float(v.value or 0) for m, v in zip(mets, r.metric_values)})
            rows.append(row)
        offset += len(res.rows)
        if offset >= res.row_count or not res.rows:
            return rows


def classify_ai(source, channel):
    s = (source or "").lower()
    for name, pat in AI_SOURCES:
        if re.search(pat, s):
            return name
    return "기타 AI" if channel == AI_CHANNEL else None


def channel_bucket(channel):
    return channel if channel in CHANNELS else OTHER_CHANNEL


def collect_week(client, property_id, start, end):
    """({콘텐츠: {컬럼: 값}}, {(콘텐츠, 채널): {컬럼: 값}}, 집계된 이벤트명 set) 반환"""
    content_rows, channel_rows, seen_events = {}, {}, set()
    for seq, name in CONTENTS.items():
        url = URL_REGEX.format(seq=seq)

        # 차원 없이 조회해야 세션·사용자가 중복 제거됨
        page = run(client, property_id, start, end, [],
                   ["sessions", "engagedSessions", "totalUsers", "activeUsers", "screenPageViews",
                    "userEngagementDuration"],
                   _filter("pagePathPlusQueryString", url))
        p = page[0] if page else defaultdict(float)

        land = run(client, property_id, start, end, ["sessionDefaultChannelGroup", "sessionSource"], ["sessions"],
                   _filter("landingPagePlusQueryString", url))

        ev = run(client, property_id, start, end, ["sessionDefaultChannelGroup", "eventName"], [CONVERSION_METRIC],
                 _and(_filter("landingPagePlusQueryString", url),
                      _in_list("eventName", TARGET_EVENTS)))

        c = defaultdict(float)
        c["세션"] = p["sessions"]
        c["참여 세션"] = p["engagedSessions"]
        c["사용자"] = p["totalUsers"]
        c["활성 사용자"] = p["activeUsers"]
        c["페이지뷰 수"] = p["screenPageViews"]
        c["평균 참여율"] = p["engagedSessions"] / p["sessions"] if p["sessions"] else 0
        c["평균 참여시간(초)"] = round(p["userEngagementDuration"] / p["activeUsers"], 1) if p["activeUsers"] else 0
        for ai_name in ("ChatGPT", "Perplexity", "Gemini", "기타 AI"):
            c[ai_name] = 0

        for r in land:
            ch, src, n = r["sessionDefaultChannelGroup"], r["sessionSource"], r["sessions"]
            c["유입 세션"] += n
            if ch == "Organic Search":
                c["오가닉 세션"] += n
            ai = classify_ai(src, ch)
            if ai:
                c[ai] += n
            channel_rows.setdefault((name, channel_bucket(ch)), defaultdict(float))["유입 세션"] += n

        for r in ev:
            en = r["eventName"]
            col = "가입시작" if en in START_EVENTS else "가입완료"
            seen_events.add(en)
            n = r[CONVERSION_METRIC]
            c[col] += n
            channel_rows.setdefault((name, channel_bucket(r["sessionDefaultChannelGroup"])), defaultdict(float))[col] += n

        content_rows[name] = {col: c[col] for col in CONTENT_COLS}

    for name in CONTENTS.values():
        for ch in CHANNELS + [OTHER_CHANNEL]:
            d = channel_rows.setdefault((name, ch), defaultdict(float))
            for col in CHANNEL_COLS:
                d[col] += 0
    return content_rows, channel_rows, seen_events


# ── 주차 ─────────────────────────────────────────────────────────────────────
def week_range(i):
    s = WEEK1_START + dt.timedelta(days=7 * (i - 1))
    return s, s + dt.timedelta(days=6)


def week_label(i):
    s, e = week_range(i)
    return f"{i:02d}주차 ({s.month}/{s.day}~{e.month}/{e.day})"


def finished_weeks(today):
    return [i for i in range(1, N_WEEKS + 1) if week_range(i)[1] < today]


def to_frames(results):
    """results: {주차번호: (content_rows, channel_rows)}"""
    content, channel = [], []
    for i in sorted(results):
        label = week_label(i)
        cdata, hdata = results[i]
        for name in CONTENTS.values():
            content.append({"주차": label, "콘텐츠": name, **{k: cdata[name].get(k) for k in CONTENT_COLS}})
            for ch in CHANNELS + [OTHER_CHANNEL]:
                channel.append({"주차": label, "콘텐츠": name, "유입채널": ch,
                                **{k: hdata[(name, ch)].get(k, 0) for k in CHANNEL_COLS}})
    return pd.DataFrame(content), pd.DataFrame(channel)


def _copy_value(col, v):
    if v is None or pd.isna(v):
        return ""
    if col == "평균 참여율":
        return f"{v * 100:.4f}%"
    return str(int(v)) if float(v).is_integer() else str(v)


def block_tsv(df, cols):
    return "\n".join("\t".join(_copy_value(c, row[c]) for c in cols) for _, row in df.iterrows())


# ── 엑셀 (XML 직접 수정: 피벗·슬라이서 보존) ─────────────────────────────────
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def new_cell(col, row, attrs, value):
    style = re.search(r'\bs="(\d+)"', attrs or "")
    s = f' s="{style.group(1)}"' if style else ""
    if value is None:
        return f'<c r="{col}{row}"{s}/>'
    v = int(value) if float(value).is_integer() else round(float(value), 6)
    return f'<c r="{col}{row}"{s}><v>{v}</v></c>'


class Workbook:
    def __init__(self, data: bytes):
        with zipfile.ZipFile(BytesIO(data)) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
        self.shared = self._shared_strings()
        self.sheet_paths = self._sheet_paths()

    def _shared_strings(self):
        data = self.files.get("xl/sharedStrings.xml")
        if not data:
            return []
        out = []
        for si in re.findall(r"<si>(.*?)</si>|<si/>", data.decode("utf-8"), re.S):
            # 윗주(rPh) 텍스트가 섞이면 라벨 매칭이 깨짐
            si = re.sub(r"<rPh\b.*?</rPh>", "", si, flags=re.S)
            out.append(html.unescape("".join(re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", si, re.S))))
        return out

    def _sheet_paths(self):
        wb = self.files["xl/workbook.xml"].decode("utf-8")
        rels = self.files["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rid2target = {}
        for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
            attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
            rid2target[attrs.get("Id")] = attrs.get("Target", "")
        out = {}
        for m in re.finditer(r"<sheet\b([^>]*)/>", wb):
            attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
            t = rid2target[attrs["r:id"]].lstrip("/")
            out[html.unescape(attrs["name"])] = t if t.startswith("xl/") else "xl/" + t
        return out

    def _cell_text(self, attrs, inner):
        inner = inner or ""
        if 't="s"' in attrs:
            m = re.search(r"<v>(\d+)</v>", inner)
            return self.shared[int(m.group(1))] if m else ""
        m = re.search(r"<t(?:\s[^>]*)?>(.*?)</t>", inner, re.S) or re.search(r"<v>(.*?)</v>", inner, re.S)
        return html.unescape(m.group(1)) if m else ""

    def find_rows(self, sheet, key_cols):
        xml = self.files[self.sheet_paths[sheet]].decode("utf-8")
        out = {}
        for rm in ROW_RE.finditer(xml):
            vals = {}
            for cm in CELL_RE.finditer(rm.group(3) or ""):
                if cm.group(1) in key_cols:
                    vals[cm.group(1)] = self._cell_text(cm.group(3), cm.group(4))
            out[tuple(vals.get(k, "") for k in key_cols)] = int(rm.group(1))
        return out

    def header(self, sheet):
        xml = self.files[self.sheet_paths[sheet]].decode("utf-8")
        rm = next(m for m in ROW_RE.finditer(xml) if m.group(1) == "1")
        return {self._cell_text(cm.group(3), cm.group(4)): cm.group(1) for cm in CELL_RE.finditer(rm.group(3) or "")}

    def write_numbers(self, sheet, values):
        path = self.sheet_paths[sheet]
        xml = self.files[path].decode("utf-8")
        by_row = defaultdict(dict)
        for (r, c), v in values.items():
            by_row[r][c] = v

        def fix_row(rm):
            r = int(rm.group(1))
            if r not in by_row:
                return rm.group(0)
            todo = dict(by_row[r])
            cells = []
            for cm in CELL_RE.finditer(rm.group(3) or ""):
                col, attrs = cm.group(1), cm.group(3)
                if col in todo:
                    cells.append((col_num(col), new_cell(col, r, attrs, todo.pop(col))))
                else:
                    cells.append((col_num(col), cm.group(0)))
            for col, v in todo.items():
                cells.append((col_num(col), new_cell(col, r, "", v)))
            cells.sort(key=lambda x: x[0])
            return f'<row r="{r}"{rm.group(2)}>' + "".join(c for _, c in cells) + "</row>"

        self.files[path] = ROW_RE.sub(fix_row, xml).encode("utf-8")

    def to_bytes(self) -> bytes:
        wb = self.files["xl/workbook.xml"].decode("utf-8")
        if "<calcPr" in wb:
            if "fullCalcOnLoad" not in wb:
                wb = wb.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
        else:
            wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self.files["xl/workbook.xml"] = wb.encode("utf-8")
        self.files.pop("xl/calcChain.xml", None)
        ct = self.files["[Content_Types].xml"].decode("utf-8")
        self.files["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct).encode("utf-8")
        rels = self.files["xl/_rels/workbook.xml.rels"].decode("utf-8")
        self.files["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", rels).encode("utf-8")
        out = BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for n in ["[Content_Types].xml"] + [n for n in self.files if n != "[Content_Types].xml"]:
                z.writestr(n, self.files[n])
        return out.getvalue()


def fill_report(xlsx: bytes, results) -> tuple[bytes, list[str]]:
    """조회한 주차의 GA4 칸만 채운 엑셀 bytes와 경고 목록 반환 (GSC 칸은 유지)"""
    wb = Workbook(xlsx)
    missing_sheets = [s for s in (SHEET_CONTENT, SHEET_CHANNEL) if s not in wb.sheet_paths]
    if missing_sheets:
        raise ValueError(f"엑셀에서 시트를 못 찾음: {missing_sheets}")
    hc, hh = wb.header(SHEET_CONTENT), wb.header(SHEET_CHANNEL)
    missing = [c for c in CONTENT_COLS if c not in hc] + [c for c in CHANNEL_COLS if c not in hh]
    if missing:
        raise ValueError(f"엑셀에서 헤더를 못 찾음: {missing}")
    rows_c = wb.find_rows(SHEET_CONTENT, ["A", "B"])
    rows_h = wb.find_rows(SHEET_CHANNEL, ["A", "B", "C"])

    warnings, out_c, out_h = [], {}, {}
    for i, (cdata, hdata) in results.items():
        label = week_label(i)
        for name, vals in cdata.items():
            r = rows_c.get((label, name))
            if not r:
                warnings.append(f"'{SHEET_CONTENT}'에 행 없음: {label} / {name}")
                continue
            for col in CONTENT_COLS:
                out_c[(r, hc[col])] = vals.get(col)
        for (name, ch), vals in hdata.items():
            r = rows_h.get((label, name, ch))
            if not r:
                warnings.append(f"'{SHEET_CHANNEL}'에 행 없음: {label} / {name} / {ch}")
                continue
            for col in CHANNEL_COLS:
                out_h[(r, hh[col])] = vals.get(col, 0)

    wb.write_numbers(SHEET_CONTENT, out_c)
    wb.write_numbers(SHEET_CHANNEL, out_h)
    return wb.to_bytes(), warnings
