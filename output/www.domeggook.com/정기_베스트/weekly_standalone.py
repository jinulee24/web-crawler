"""도매꾹 베스트 정기 수집 — 메일 전용판 (Claude/MCP/노션 없이 동작).

이 파일 하나로 아래 3단계를 수행한다:
  1) crawl_script.py 실행 (수집 + 엑셀 2개 저장)
  2) Claude API로 분석 리포트(Brief) 작성 (ANTHROPIC_API_KEY 없으면 숫자 템플릿)
  3) Gmail(SMTP)로 분석 리포트 본문 + 엑셀 2개 첨부를 발송

노션은 쓰지 않는다.

── 필요한 것 ──────────────────────────────────────────────────────────────
  pip install -r requirements.txt   (레포 루트에서, 최초 1회)
  pip install requests              (이 스크립트 전용 — requirements.txt 엔 없음)

  환경변수 (또는 이 파일과 같은 폴더의 weekly.env 파일에 KEY=VALUE 로 적어도 됨):
    GMAIL_ADDRESS         필수. 보내는 계정 (예: jinu.lee24@gmail.com)
    GMAIL_APP_PASSWORD    필수. Google 계정 → 보안 → 2단계 인증 켠 뒤 "앱 비밀번호" 발급 (16자리).
    MAIL_TO               선택. 받는 사람. 콤마로 여러 명 가능. 기본값 jinu.lee24@gmail.com
    ANTHROPIC_API_KEY     선택. 있으면 Claude API로 Brief를 실제로 써서 채운다.
    PYTHON_EXE            선택. crawl_script.py 를 실행할 python 경로. 기본값은 이 레포의 .venv.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import smtplib
import subprocess
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
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


# ── 3. Gmail 발송 (본문 + 엑셀 첨부) ──────────────────────────────────────

def send_gmail(subject: str, html_body: str, text_body: str, attachments: list[Path] | None = None):
    addr = env("GMAIL_ADDRESS")
    app_pw = env("GMAIL_APP_PASSWORD")
    # MAIL_TO 는 콤마로 여러 주소를 넣을 수 있다 (예: "a@x.com, b@y.com")
    to_list = [t.strip() for t in os.environ.get("MAIL_TO", "jinu.lee24@gmail.com").split(",") if t.strip()]

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = addr
    msg["To"] = ", ".join(to_list)

    body = MIMEMultipart("alternative")
    body.attach(MIMEText(text_body, "plain", "utf-8"))
    body.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(body)

    for path in attachments or []:
        part = MIMEApplication(Path(path).read_bytes(),
                               _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # 한글 파일명은 RFC 2231 로 인코딩해야 메일 클라이언트에서 안 깨진다
        part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", Path(path).name))
        msg.attach(part)

    print(f"[3/3] 메일 발송: {', '.join(t[:2] + '***' for t in to_list)} (첨부 {len(attachments or [])}개)")
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
        s.starttls()
        s.login(addr, app_pw)
        s.sendmail(addr, to_list, msg.as_string())


def html_table(headers, rows):
    """rows 의 셀은 이미 HTML 로 안전하게 만들어진 문자열(또는 숫자)이어야 한다."""
    th = "".join(f"<th style='border:1px solid #ccc;padding:4px 8px;background:#f3f3f3'>{h}</th>" for h in headers)
    trs = []
    for row in rows:
        tds = "".join(f"<td style='border:1px solid #ccc;padding:4px 8px'>{c}</td>" for c in row)
        trs.append(f"<tr>{tds}</tr>")
    return f"<table style='border-collapse:collapse;font-size:13px'><tr>{th}</tr>{''.join(trs)}</table>"


def build_report(summary: dict, brief: dict, now: datetime) -> tuple[str, str, str]:
    """(제목, HTML 본문, 평문 본문) — 분석 리포트."""
    weekday = WEEKDAY_KO[now.weekday()]
    rising_rows = summary["brief_candidates_rising"]
    new_rows = summary["brief_candidates_new"]
    e = escape

    subject = f"[도매꾹 베스트] {now.strftime('%Y-%m-%d')} ({weekday}) — {brief['one_liner']}"
    overview = (f"수집 {summary['run_at']} · 5개 카테고리(가전 제외) × 1~150위 · 전체 {summary['total']}건 · "
                f"급상승(▲{RISE_THRESHOLD}+) {summary['rise10_total']}건 · NEW {summary['new_total']}건 · "
                f"실패 카테고리 {', '.join(summary['failed_categories']) or '없음'}")

    cat_rows = [[e(c), v["total"], v["rise10"], v["new"], v["price_hidden"]]
                for c, v in summary["per_category"].items()]
    warnings = (summary.get("validation_issues") or []) + (summary.get("pii_warnings") or [])

    html = (
        f"<h2>도매꾹 베스트 주간 리포트</h2><p>{e(overview)}</p>"
        f"<h3>이번 회차 한눈에</h3><ul>{''.join(f'<li>{e(str(b))}</li>' for b in brief['highlights'])}</ul>"
        f"<h3>카테고리별 현황</h3>"
        + html_table(["카테고리", "수집", f"▲{RISE_THRESHOLD}+", "NEW", "가격비공개"], cat_rows)
        + "<h3>급상승 TOP 30</h3>"
        + html_table(
            ["#", "카테고리", "순위", "변동", "상품명", "판매가"],
            [[i, e(r["카테고리명"]), r["랭킹"], e(r["순위변동"]),
              f"<a href='{e(r['상품URL'])}'>{e(clean_name(r['상품명']))}</a>", price_str(r["판매가격"])]
             for i, r in enumerate(rising_rows, 1)])
        + "<h3>NEW 진입 TOP 30</h3>"
        + html_table(
            ["#", "카테고리", "진입순위", "상품명", "판매가"],
            [[i, e(r["카테고리명"]), r["랭킹"],
              f"<a href='{e(r['상품URL'])}'>{e(clean_name(r['상품명']))}</a>", price_str(r["판매가격"])]
             for i, r in enumerate(new_rows, 1)])
        + f"<h3>시사점</h3><ul>{''.join(f'<li>{e(str(b))}</li>' for b in brief['insights'])}</ul>"
        + (f"<h3>검증 경고</h3><ul>{''.join(f'<li>{e(str(w))}</li>' for w in warnings)}</ul>" if warnings else "")
        + "<p>📎 첨부: 전체 목록 엑셀(카테고리별 시트 포함) · 급상승(▲10+)/NEW 엑셀</p>"
    )

    lines = [overview, "", "[이번 회차 한눈에]"] + [f"- {b}" for b in brief["highlights"]]
    lines += ["", "[시사점]"] + [f"- {b}" for b in brief["insights"]]
    if warnings:
        lines += ["", "[검증 경고]"] + [f"- {w}" for w in warnings]
    lines += ["", "첨부: 전체 목록 엑셀 · 급상승/NEW 엑셀", "(상세 표는 HTML 메일 본문과 엑셀을 확인하세요)"]
    return subject, html, "\n".join(lines)


# ── main ──────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="수집·분석은 실제로 하되 메일 발송 없이 내용만 출력")
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
                f"<p>{escape(str(e))}</p><p>우회는 시도하지 않았음.</p>",
                f"{e}\n우회는 시도하지 않았음.",
            )
        raise

    summary = json.loads((run_dir / "summary.json").read_text(encoding="utf-8"))
    brief = build_brief(summary)
    attachments = [Path(summary["all_xlsx"]), Path(summary["highlight_xlsx"])]
    subject, html, text = build_report(summary, brief, datetime.now())

    if args.dry_run:
        print("\n=== DRY RUN — 메일 발송 생략, 내용만 출력 ===")
        print("제목:", subject)
        print("첨부 예정:", ", ".join(p.name for p in attachments))
        print(json.dumps(brief, ensure_ascii=False, indent=2))
        return

    send_gmail(subject, html, text, attachments)

    print(f"\n완료 — RUN_DIR={run_dir}")
    print(f"총 {summary['total']}건 / 급상승 {summary['rise10_total']} / NEW {summary['new_total']}")
    print("메일: 발송됨 (본문 리포트 + 엑셀 2개 첨부)")


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
# 공통: weekly.env 파일에 GMAIL_ADDRESS / GMAIL_APP_PASSWORD 를 적어 두면
#       이 스크립트가 실행 시점에 자동으로 읽는다 (systemd는 EnvironmentFile= 로 직접 지정).

