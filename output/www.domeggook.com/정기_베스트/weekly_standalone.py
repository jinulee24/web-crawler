"""도매꾹 베스트 정기 수집 — 서버 단독 실행판 (Claude/MCP 없이 동작).

이 파일 하나로 지금까지 Claude 세션 안에서 하던 3단계를 그대로 재현한다:
  1) crawl_script.py 실행 (수집 + 엑셀 2개 저장)  ← 기존 스크립트 그대로 재사용, 수정 없음
  2) 노션 "🛒 도매꾹 Weekly Best" DB에 새 행 추가 (Notion REST API 직접 호출, 파일 2개 첨부)
  3) Gmail 요약 메일 발송 (SMTP, 앱 비밀번호 사용)

Claude Code 앱이 켜져 있을 필요가 없다. Python + 아래 환경변수만 있으면 어떤 서버(리눅스/윈도우)에서든
cron / systemd timer / Task Scheduler 로 매주 월요일 08:00에 그대로 돌릴 수 있다.

── 필요한 것 ──────────────────────────────────────────────────────────────
  pip install -r requirements.txt   (레포 루트에서, 최초 1회)
  pip install requests              (이 스크립트 전용 — requirements.txt 엔 없음)

  환경변수 (또는 이 파일과 같은 폴더의 weekly.env 파일에 KEY=VALUE 로 적어도 됨):
    NOTION_TOKEN          필수. Notion 내부 통합(Internal Integration) 시크릿.
                          https://www.notion.so/my-integrations 에서 발급 후,
                          "🛒 도매꾹 Weekly Best" DB 페이지 ••• → 연결(Connections) 에 그 통합을 추가해야 함.
                          (Claude 쪽 Notion 커넥터와는 별개 — 이 토큰은 이 스크립트 전용으로 새로 만든다.)
    GMAIL_ADDRESS         필수. 보내는 계정 (예: jinu.lee24@gmail.com)
    GMAIL_APP_PASSWORD    필수. Google 계정 → 보안 → 2단계 인증 켠 뒤 "앱 비밀번호" 발급 (16자리).
                          일반 로그인 비밀번호 아님.
    MAIL_TO               선택. 기본값 jinu.lee24@gmail.com
    ANTHROPIC_API_KEY     선택. 있으면 Claude API로 Brief(한줄요약/한눈에/시사점)를 실제로 "써서" 채운다.
                          없으면 숫자 기반 템플릿 문장으로 대체한다 (품질은 낮지만 무료·의존성 없음).
    PYTHON_EXE             선택. crawl_script.py 를 실행할 python 경로. 기본값은 이 레포의 .venv.

  스케줄 등록 예시는 파일 맨 아래 주석 참고.
"""
from __future__ import annotations

import argparse
import json
import mimetypes
import os
import re
import smtplib
import subprocess
import sys
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("requests 패키지가 없습니다. 먼저 `pip install requests` 를 실행하세요.")

# ── 경로 ──────────────────────────────────────────────────────────────────
HERE = Path(__file__).resolve().parent                       # .../output/www.domeggook.com/정기_베스트
REPO = HERE.parents[2]                                        # 레포 루트
CRAWL_SCRIPT = HERE / "crawl_script.py"

# ── 설정 ──────────────────────────────────────────────────────────────────
NOTION_VERSION = os.environ.get("NOTION_VERSION", "2026-03-11")
NOTION_DATA_SOURCE_ID = os.environ.get(
    "NOTION_DATA_SOURCE_ID", "bf5a668f-4dee-4545-a06f-e6d3ffbff448"
)
NOTION_API = "https://api.notion.com/v1"
WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]
RISE_THRESHOLD = 10


def _load_dotenv():
    """같은 폴더의 weekly.env 가 있으면 읽어서 os.environ 에 채워 넣는다 (이미 있는 값은 안 덮음)."""
    envfile = HERE / "weekly.env"
    if not envfile.exists():
        return
    for line in envfile.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def env(name, required=True, default=None):
    val = os.environ.get(name, default)
    if required and not val:
        sys.exit(f"[설정 누락] 환경변수 {name} 가 필요합니다. weekly.env 또는 시스템 환경변수로 설정하세요.")
    return val


# ── 1. 수집 실행 ─────────────────────────────────────────────────────────

def run_crawl() -> Path:
    py = os.environ.get("PYTHON_EXE") or str(
        REPO / (".venv/Scripts/python.exe" if os.name == "nt" else ".venv/bin/python")
    )
    print(f"[1/3] 수집 실행: {py} {CRAWL_SCRIPT}")
    proc = subprocess.run(
        [py, str(CRAWL_SCRIPT)],
        cwd=str(REPO),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    print(proc.stdout)
    if proc.returncode != 0:
        print(proc.stderr, file=sys.stderr)
        raise SystemExit(
            f"crawl_script.py 실패 (exit={proc.returncode}). "
            "차단(소프트블록) 감지일 수 있음 — 우회하지 않고 여기서 중단한다."
        )
    m = re.search(r"^RUN_DIR=(.+)$", proc.stdout, re.MULTILINE)
    if not m:
        raise SystemExit("RUN_DIR 출력을 못 찾음 — crawl_script.py 출력 형식이 바뀐 것으로 보임.")
    return Path(m.group(1).strip())


# ── 2. Brief 작성 ────────────────────────────────────────────────────────

def clean_name(name: str, n: int = 42) -> str:
    """앞쪽 [태그] 제거 + 남은 대괄호 이스케이프 + 길이 제한 (마크다운 링크 파손 방지)."""
    name = re.sub(r"\s+", " ", name).strip()
    while True:
        m = re.match(r"^\[[^\[\]]{1,20}\]\s*", name)
        if not m:
            break
        name = name[m.end():]
    if len(name) > n:
        name = name[:n].rstrip() + "…"
    return name


def price_str(p) -> str:
    return f"{p:,}원" if p else "사업자전용"


def build_brief_template(summary: dict) -> dict:
    """ANTHROPIC_API_KEY 가 없을 때 쓰는 규칙 기반 Brief (숫자 요약만, 해석 문장 없음)."""
    total, rise, new = summary["total"], summary["rise10_total"], summary["new_total"]
    gift = summary.get("gift_keyword_count", 0)
    one_liner = f"급상승(▲10+) {rise}건 · NEW {new}건 · 선물세트류 {gift}건"
    return {
        "one_liner": one_liner,
        "highlights": [
            f"전체 {total}건 중 급상승(▲10 이상) {rise}건, NEW {new}건이 확인됐습니다.",
            f"선물세트·명절 키워드 포함 상품이 {gift}건입니다.",
            "카테고리별 세부 수치는 아래 표를 참고하세요.",
        ],
        "insights": [
            "이 Brief는 ANTHROPIC_API_KEY 가 설정되지 않아 템플릿으로 생성됐습니다 — "
            "해석 문장 없이 숫자만 담았습니다. 상세 해석이 필요하면 첨부 엑셀을 직접 검토하세요.",
        ],
    }


def build_brief_claude(summary: dict, api_key: str) -> dict:
    """ANTHROPIC_API_KEY 가 있으면 Claude API로 Brief 문장을 실제로 작성한다."""
    payload = {
        "categories": summary["categories"],
        "total": summary["total"],
        "rise10_total": summary["rise10_total"],
        "new_total": summary["new_total"],
        "per_category": summary["per_category"],
        "gift_keyword_count": summary.get("gift_keyword_count", 0),
        "brief_candidates_rising": summary["brief_candidates_rising"],
        "brief_candidates_new": summary["brief_candidates_new"],
    }
    prompt = (
        "다음은 도매꾹 베스트 상품 수집 결과 JSON이다. 이 데이터에서 실제로 읽히는 사실과 "
        "그에 근거한 해석만 써라. 판매량·매출처럼 데이터에 없는 수치는 지어내지 말고, "
        "해석은 '~로 보입니다'처럼 해석임이 드러나게 써라.\n\n"
        f"{json.dumps(payload, ensure_ascii=False)}\n\n"
        "아래 JSON 형식으로만 답하라 (다른 텍스트 없이 ```json 코드블록 하나):\n"
        '{"one_liner": "흐름 2~3개를 · 로 이은 한 줄",\n'
        ' "highlights": ["이번 회차 한눈에 불릿 3~4개"],\n'
        ' "insights": ["시사점 불릿 3~4개"]}'
    )
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": "claude-sonnet-5",
            "max_tokens": 4096,
            "messages": [{"role": "user", "content": prompt}],
        },
        timeout=60,
    )
    resp.raise_for_status()
    body = resp.json()
    # 이 모델은 기본적으로 thinking 블록을 먼저 내놓는다 — content[0]이 아니라
    # type이 "text"인 블록을 찾는다 (없으면 잘림/오류이므로 그대로 예외).
    text_blocks = [b["text"] for b in body.get("content", []) if b.get("type") == "text"]
    if not text_blocks:
        raise ValueError(f"텍스트 블록 없음 (stop_reason={body.get('stop_reason')}): {body}")
    text = "".join(text_blocks)
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        raise ValueError(f"Claude 응답에서 JSON을 못 찾음: {text[:200]}")
    return json.loads(m.group(0))


def build_brief(summary: dict) -> dict:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if api_key:
        try:
            print("[2/3] Brief 작성: Claude API 사용")
            return build_brief_claude(summary, api_key)
        except Exception as e:
            print(f"  Claude API 실패({e}) — 템플릿으로 대체")
    print("[2/3] Brief 작성: 템플릿(ANTHROPIC_API_KEY 없음)")
    return build_brief_template(summary)


# ── Notion REST API ──────────────────────────────────────────────────────

def notion_headers(token, json_body=True):
    h = {"Authorization": f"Bearer {token}", "Notion-Version": NOTION_VERSION}
    if json_body:
        h["Content-Type"] = "application/json"
    return h


def notion_upload_file(token: str, path: Path) -> str:
    """파일 업로드 3단계: 슬롯 생성 → 바이트 전송 → file_upload id 반환."""
    ctype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    r = requests.post(
        f"{NOTION_API}/file_uploads",
        headers=notion_headers(token),
        json={"mode": "single_part", "filename": path.name, "content_type": ctype},
        timeout=30,
    )
    r.raise_for_status()
    upload = r.json()
    upload_id, upload_url = upload["id"], upload["upload_url"]

    with open(path, "rb") as f:
        r2 = requests.post(
            upload_url,
            headers=notion_headers(token, json_body=False),
            files={"file": (path.name, f, ctype)},
            timeout=120,
        )
    r2.raise_for_status()
    status = r2.json().get("status")
    if status != "uploaded":
        raise RuntimeError(f"업로드 상태 이상: {r2.json()}")
    return upload_id


def rt(text, url=None, bold=False):
    span = {"type": "text", "text": {"content": str(text)}}  # 방어: 숫자/None 등 비-문자열이 와도 안전하게
    if url:
        span["text"]["link"] = {"url": url}
    ann = {}
    if bold:
        ann["bold"] = True
    if ann:
        span["annotations"] = ann
    return span


def heading2(text):
    return {"object": "block", "type": "heading_2",
            "heading_2": {"rich_text": [rt(text)]}}


def paragraph(text):
    return {"object": "block", "type": "paragraph",
            "paragraph": {"rich_text": [rt(text)]}}


def bulleted(text):
    return {"object": "block", "type": "bulleted_list_item",
            "bulleted_list_item": {"rich_text": [rt(text)]}}


def file_block(file_upload_id):
    return {"object": "block", "type": "file",
            "file": {"type": "file_upload", "file_upload": {"id": file_upload_id}}}


def table_block(headers, rows):
    """rows: list[list[(text, url_or_None)]] — 각 셀 (표시문자열, 링크 or None)."""
    def row_block(cells):
        return {"object": "block", "type": "table_row",
                "table_row": {"cells": [[rt(txt, url)] for txt, url in cells]}}

    header_row = row_block([(h, None) for h in headers])
    body_rows = [row_block(r) for r in rows]
    return {
        "object": "block",
        "type": "table",
        "table": {
            "table_width": len(headers),
            "has_column_header": True,
            "has_row_header": False,
            "children": [header_row] + body_rows,
        },
    }


def notion_create_page(token: str, properties: dict, first_children: list) -> tuple[str, str]:
    """스키마에 없는 속성은 하나씩 빼면서 재시도 — DB 스키마를 절대 바꾸지 않는다."""
    props = dict(properties)
    for _ in range(len(props) + 1):
        body = {
            "parent": {"type": "data_source_id", "data_source_id": NOTION_DATA_SOURCE_ID},
            "properties": props,
            "children": first_children,
        }
        r = requests.post(f"{NOTION_API}/pages", headers=notion_headers(token), json=body, timeout=30)
        if r.status_code < 300:
            data = r.json()
            return data["id"], data["url"]
        try:
            msg = r.json().get("message", "")
        except Exception:
            msg = r.text
        bad = None
        for name in list(props.keys()):
            if name in msg:
                bad = name
                break
        if bad:
            print(f"  속성 '{bad}' 를 스키마가 안 받아 제외하고 재시도: {msg[:150]}")
            del props[bad]
            continue
        r.raise_for_status()
    raise RuntimeError("페이지 생성 실패 — 속성을 다 빼도 안 됨")


def notion_append(token: str, page_id: str, blocks: list):
    # Notion append 는 한 번에 최대 100블록 — 안전하게 40개씩 나눠 보낸다.
    for i in range(0, len(blocks), 40):
        chunk = blocks[i:i + 40]
        r = requests.patch(
            f"{NOTION_API}/blocks/{page_id}/children",
            headers=notion_headers(token),
            json={"children": chunk},
            timeout=30,
        )
        if r.status_code >= 300:
            # 원인을 반드시 화면에 남긴다 — 이게 없으면 "400 Bad Request"라는 것만 알고 끝난다.
            print(f"  [notion_append 실패] status={r.status_code} body={r.text[:1500]}")
        r.raise_for_status()
        time.sleep(0.3)


def verify_page(token: str, page_id: str, expect_files: int, expect_tables: int) -> list[str]:
    """등록 직후 실제로 반영됐는지 재조회로 확인한다. 문제 목록(빈 리스트=정상)을 반환."""
    problems = []
    r = requests.get(f"{NOTION_API}/blocks/{page_id}/children?page_size=100",
                      headers=notion_headers(token), timeout=30)
    r.raise_for_status()
    blocks = r.json().get("results", [])
    n_files = sum(1 for b in blocks if b.get("type") == "file")
    n_tables = sum(1 for b in blocks if b.get("type") == "table")
    if n_files < expect_files:
        problems.append(f"첨부 파일 {n_files}/{expect_files}개만 확인됨")
    if n_tables < expect_tables:
        problems.append(f"표 블록 {n_tables}/{expect_tables}개만 확인됨")
    return problems


def build_notion_children(brief: dict, summary: dict, rising_rows, new_rows) -> list[list]:
    """append 호출 단위로 나눠서 반환 (파일 첨부부터 순서대로)."""
    cat_headers = ["카테고리", "수집", f"▲{RISE_THRESHOLD}+", "NEW", "가격비공개"]
    cat_rows = [
        [(cat, None), (str(v["total"]), None), (str(v["rise10"]), None),
         (str(v["new"]), None), (str(v["price_hidden"]), None)]
        for cat, v in summary["per_category"].items()
    ]
    cat_rows.append([
        ("합계", None), (str(summary["total"]), None), (str(summary["rise10_total"]), None),
        (str(summary["new_total"]), None),
        (str(sum(v["price_hidden"] for v in summary["per_category"].values())), None),
    ])

    rising_headers = ["#", "카테고리", "순위", "변동", "상품명", "판매가"]
    rising_table_rows = [
        [(str(i), None), (r["카테고리명"], None), (str(r["랭킹"]), None), (r["순위변동"], None),
         (clean_name(r["상품명"]), r["상품URL"]), (price_str(r["판매가격"]), None)]
        for i, r in enumerate(rising_rows, 1)
    ]

    new_headers = ["#", "카테고리", "진입순위", "상품명", "판매가"]
    new_table_rows = [
        [(str(i), None), (r["카테고리명"], None), (str(r["랭킹"]), None),
         (clean_name(r["상품명"]), r["상품URL"]), (price_str(r["판매가격"]), None)]
        for i, r in enumerate(new_rows, 1)
    ]

    chunks = []
    chunks.append(["__FILES__"])  # placeholder — 실제 file_upload id는 main()에서 채운다
    chunks.append(
        [heading2("1. Executive Summary")] + [bulleted(b) for b in brief["highlights"]]
    )
    chunks.append([heading2("2. 카테고리별 현황"), table_block(cat_headers, cat_rows)])
    chunks.append([heading2("3. 급상승 TOP 30"), table_block(rising_headers, rising_table_rows)])
    chunks.append([heading2("4. NEW 진입 TOP 30"), table_block(new_headers, new_table_rows)])
    chunks.append([heading2("5. 시사점")] + [bulleted(b) for b in brief["insights"]])
    issues = summary.get("validation_issues") or []
    pii = summary.get("pii_warnings") or []
    if issues or pii:
        chunks.append([heading2("6. 검증 경고")] + [bulleted(x) for x in (issues + pii)])
    return chunks


# ── 3. Gmail 발송 ────────────────────────────────────────────────────────

def send_gmail(subject: str, html_body: str, text_body: str, success: bool):
    addr = env("GMAIL_ADDRESS")
    app_pw = env("GMAIL_APP_PASSWORD")
    to = os.environ.get("MAIL_TO", "jinu.lee24@gmail.com")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = addr
    msg["To"] = to
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    print(f"[3/3] 메일 발송: {to}")
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
        s.starttls()
        s.login(addr, app_pw)
        s.sendmail(addr, [to], msg.as_string())


def html_table(headers, rows):
    th = "".join(f"<th style='border:1px solid #ccc;padding:4px 8px'>{h}</th>" for h in headers)
    trs = []
    for row in rows:
        tds = "".join(f"<td style='border:1px solid #ccc;padding:4px 8px'>{c}</td>" for c in row)
        trs.append(f"<tr>{tds}</tr>")
    return f"<table style='border-collapse:collapse'><tr>{th}</tr>{''.join(trs)}</table>"


# ── main ──────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                     help="수집은 실제로 하되 노션 등록·메일 발송 없이 내용만 출력")
    args = ap.parse_args()

    _load_dotenv()

    try:
        run_dir = run_crawl()
    except SystemExit as e:
        # 수집 실패 — 우회하지 않고 실패 메일만 보낸다
        today = datetime.now().strftime("%Y-%m-%d")
        if not args.dry_run:
            send_gmail(
                f"[도매꾹 베스트] {today} 수집 실패",
                f"<p>{e}</p><p>우회는 시도하지 않았음.</p>",
                f"{e}\n우회는 시도하지 않았음.",
                success=False,
            )
        raise

    summary = json.loads((run_dir / "summary.json").read_text(encoding="utf-8"))
    brief = build_brief(summary)

    now = datetime.now()
    weekday = WEEKDAY_KO[now.weekday()]
    title = f"{now.strftime('%Y-%m-%d')} ({weekday}) 도매꾹 베스트"
    reason = (f"{brief['one_liner']} — 전체 {summary['total']} · "
              f"급상승(▲{RISE_THRESHOLD}+) {summary['rise10_total']} · NEW {summary['new_total']}")

    properties = {
        "상품명": {"title": [rt(title)]},
        "추천 이유": {"rich_text": [rt(reason)]},
        "링크": {"url": "https://www.domeggook.com/main/item/itemPopular.php?cat=01_00"},
        "date:큐레이션 주차:start": now.strftime("%Y-%m-%d"),
        "date:큐레이션 주차:is_datetime": 0,
    }
    # notion_create_page 는 dict of {name: {...notion property value...}} 를 그대로 pages API에 보낸다.
    # 위에서 date: 접두 키는 편의상 우리 내부 표기이므로 실제 전송 전에 변환한다.
    def to_notion_props(props):
        out = {}
        date_val = {}
        for k, v in props.items():
            if k.startswith("date:"):
                _, field, part = k.split(":")
                date_val.setdefault(field, {})[part] = v
            else:
                out[k] = v
        for field, parts in date_val.items():
            out[field] = {"date": {"start": parts.get("start"),
                                    "end": None}}
            # is_datetime 은 Notion Pages API 표준 필드가 아니라 표시 방식 문제라 별도 처리 불필요.
        return out

    all_xlsx = Path(summary["all_xlsx"])
    highlight_xlsx = Path(summary["highlight_xlsx"])
    rising_rows = summary["brief_candidates_rising"]
    new_rows = summary["brief_candidates_new"]
    chunks = build_notion_children(brief, summary, rising_rows, new_rows)

    if args.dry_run:
        print("\n=== DRY RUN — 노션/메일 생략, 내용만 출력 ===")
        print("제목:", title)
        print("추천 이유:", reason)
        print("업로드 예정 파일:", all_xlsx.name, highlight_xlsx.name)
        print(json.dumps(brief, ensure_ascii=False, indent=2))
        return

    token = env("NOTION_TOKEN")
    print("[2/3] 노션 파일 업로드")
    file_id_all = notion_upload_file(token, all_xlsx)
    file_id_hl = notion_upload_file(token, highlight_xlsx)

    first_chunk_idx = next(i for i, c in enumerate(chunks) if c == ["__FILES__"])
    chunks[first_chunk_idx] = [
        heading2("📎 이번 회차 파일"),
        file_block(file_id_all),
        file_block(file_id_hl),
        paragraph(
            f"수집 {summary['run_at']} · 5개 카테고리(가전 제외) × 1~150위 · "
            f"전체 {summary['total']}건 · 급상승(▲{RISE_THRESHOLD} 이상) {summary['rise10_total']}건 · "
            f"NEW {summary['new_total']}건 · 실패 카테고리 "
            f"{', '.join(summary['failed_categories']) or '없음'}"
        ),
    ]

    print("[2/3] 노션 페이지 생성")
    page_id, page_url = notion_create_page(token, to_notion_props(properties), chunks[0])
    append_failures = []
    for idx, chunk in enumerate(chunks[1:], start=1):
        try:
            notion_append(token, page_id, chunk)
        except Exception as e:
            # 본문 한 덩어리가 실패해도 페이지 자체(+파일 첨부)는 이미 살아있다 —
            # 나머지 덩어리는 계속 시도하고, 실패한 것만 기록해 끝까지 보고한다.
            print(f"  [청크 {idx}/{len(chunks)-1} 추가 실패, 계속 진행] {e}")
            append_failures.append(idx)

    problems = verify_page(token, page_id, expect_files=2, expect_tables=3)
    if append_failures:
        problems.append(f"본문 청크 {append_failures} 추가 실패 (위 로그의 [notion_append 실패] 참고)")
    if problems:
        print("  ⚠ 검증 경고:", "; ".join(problems))
    else:
        print("  ✓ 재조회 확인: 파일 2개·표 3개 정상 반영")

    subject = f"[도매꾹 베스트] {now.strftime('%Y-%m-%d')} ({weekday}) — {brief['one_liner']}"
    html = (
        f"<p>{reason}</p>"
        f"<h3>이번 회차 한눈에</h3><ul>{''.join(f'<li>{b}</li>' for b in brief['highlights'])}</ul>"
        f"<h3>급상승 TOP 30</h3>"
        + html_table(
            ["#", "카테고리", "순위", "변동", "상품명", "판매가"],
            [[i, r["카테고리명"], r["랭킹"], r["순위변동"],
              f"<a href='{r['상품URL']}'>{clean_name(r['상품명'])}</a>", price_str(r["판매가격"])]
             for i, r in enumerate(rising_rows, 1)],
        )
        + "<h3>NEW 진입 TOP 30</h3>"
        + html_table(
            ["#", "카테고리", "진입순위", "상품명", "판매가"],
            [[i, r["카테고리명"], r["랭킹"],
              f"<a href='{r['상품URL']}'>{clean_name(r['상품명'])}</a>", price_str(r["판매가격"])]
             for i, r in enumerate(new_rows, 1)],
        )
        + f"<h3>시사점</h3><ul>{''.join(f'<li>{b}</li>' for b in brief['insights'])}</ul>"
        + f"<p>노션: <a href='{page_url}'>{page_url}</a></p>"
        + f"<p>로컬 파일: {all_xlsx}<br>{highlight_xlsx}</p>"
    )
    text = (
        f"{reason}\n\n노션: {page_url}\n로컬 파일:\n{all_xlsx}\n{highlight_xlsx}\n"
    )
    send_gmail(subject, html, text, success=True)

    print(f"\n완료 — RUN_DIR={run_dir}")
    print(f"총 {summary['total']}건 / 급상승 {summary['rise10_total']} / NEW {summary['new_total']}")
    print(f"노션: {page_url}")
    print("메일: 발송됨" if not problems else f"메일: 발송됨 (노션 검증 경고 있음: {problems})")


if __name__ == "__main__":
    main()


# ═══════════════════════════════════════════════════════════════════════
# 스케줄 등록 예시 (매주 월요일 08:00)
# ═══════════════════════════════════════════════════════════════════════
#
# ── Linux/macOS 서버: cron ──
#   crontab -e 로 아래 한 줄 추가 (경로는 실제 배포 위치로 바꿀 것):
#
#   0 8 * * 1 cd /opt/web-crawler && .venv/bin/python \
#     output/www.domeggook.com/정기_베스트/weekly_standalone.py \
#     >> /opt/web-crawler/logs/domeggook_weekly.log 2>&1
#
# ── Linux 서버: systemd timer (cron보다 로그·재시도 관리가 편함) ──
#   /etc/systemd/system/domeggook-weekly.service
#     [Unit]
#     Description=도매꾹 베스트 정기 수집
#     [Service]
#     Type=oneshot
#     WorkingDirectory=/opt/web-crawler
#     ExecStart=/opt/web-crawler/.venv/bin/python output/www.domeggook.com/정기_베스트/weekly_standalone.py
#     EnvironmentFile=/opt/web-crawler/output/www.domeggook.com/정기_베스트/weekly.env
#
#   /etc/systemd/system/domeggook-weekly.timer
#     [Unit]
#     Description=매주 월요일 08:00 도매꾹 베스트
#     [Timer]
#     OnCalendar=Mon *-*-* 08:00:00
#     Persistent=true   # ← 서버가 그 시간에 꺼져 있었어도 켜지면 곧바로 실행
#     [Install]
#     WantedBy=timers.target
#
#   등록: sudo systemctl enable --now domeggook-weekly.timer
#
# ── Windows 서버: 작업 스케줄러 ──
#   schtasks /Create /TN "도매꾹베스트" /SC WEEKLY /D MON /ST 08:00 ^
#     /TR "C:\web-crawler\.venv\Scripts\python.exe C:\web-crawler\output\www.domeggook.com\정기_베스트\weekly_standalone.py" ^
#     /RL HIGHEST
#   → "사용자 로그온 여부와 관계없이 실행" 체크(속성 창)해야 컴퓨터에 아무도 로그인 안 해도 돈다.
#
# 공통: weekly.env 파일에 NOTION_TOKEN / GMAIL_ADDRESS / GMAIL_APP_PASSWORD 를 적어 두면
#       이 스크립트가 실행 시점에 자동으로 읽는다 (systemd는 EnvironmentFile= 로 직접 지정).
