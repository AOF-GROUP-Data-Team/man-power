name: Manpower report

on:
  workflow_dispatch:
    inputs:
      mode:
        description: "auto = send only on Saturday (KSA) | force = send now | test = build xlsx, no email"
        required: false
        default: "auto"
        type: choice
        options: ["auto", "force", "test"]

jobs:
  manpower:
    runs-on: ubuntu-latest
    timeout-minutes: 30
    env:
      TZ: Asia/Riyadh
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"
          cache: pip

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Build and send manpower report
        env:
          ZENPUT_API_TOKEN: ${{ secrets.ZENPUT_API_TOKEN }}
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
          MODE: ${{ github.event.inputs.mode || 'auto' }}
        run: python manpower_report.py

      - name: Upload xlsx
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: manpower-report
          path: manpower_report_*.xlsx
          if-no-files-found: ignore
          retention-days: 14
