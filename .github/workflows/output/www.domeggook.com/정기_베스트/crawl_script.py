name: 도매꾹 베스트 정기 수집

on:
  schedule:
    # 매주 월요일 08:00 KST = 일요일 23:00 UTC
    - cron: "0 23 * * 0"
  workflow_dispatch: {}   # 수동으로 "Run workflow" 버튼으로도 실행 가능 (테스트용)

jobs:
  run:
    runs-on: ubuntu-latest
    timeout-minutes: 15
    steps:
      - name: 레포 체크아웃
        uses: actions/checkout@v4

      - name: Python 설치
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: 의존성 설치
        run: |
          pip install -r requirements.txt
          pip install requests

      - name: 정기 수집 실행
        working-directory: output/www.domeggook.com/정기_베스트
        env:
          PYTHON_EXE: python           # GitHub Actions엔 .venv가 없다 — 시스템 python을 직접 쓴다
          NOTION_TOKEN: ${{ secrets.NOTION_TOKEN }}
          GMAIL_ADDRESS: ${{ secrets.GMAIL_ADDRESS }}
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
          MAIL_TO: ${{ secrets.MAIL_TO }}
          ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
        run: python weekly_standalone.py

      - name: 실패 시 로그 보존
        if: failure()
        uses: actions/upload-artifact@v4
        with:
          name: domeggook-weekly-failure-log
          path: output/www.domeggook.com/정기_베스트/*/raw_data.json
          if-no-files-found: ignore
