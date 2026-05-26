name: Update Dashboard Data

on:
  schedule:
    - cron: '0 1 * * *'   # Setiap hari 01:00 UTC = 08:00 WIB
  workflow_dispatch:        # Manual trigger dari GitHub UI

jobs:
  update-data:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout repo
        uses: actions/checkout@v4

      - name: Set up Python 3.11
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Fetch & process data from Google Sheets
        env:
          RETAIL_2025_URL:     ${{ secrets.RETAIL_2025_URL }}
          RETAIL_2026_URL:     ${{ secrets.RETAIL_2026_URL }}
          INDUSTRIAL_2025_URL: ${{ secrets.INDUSTRIAL_2025_URL }}
          INDUSTRIAL_2026_URL: ${{ secrets.INDUSTRIAL_2026_URL }}
        run: python scripts/fetch_data.py

      - name: Commit & push if data changed
        run: |
          git config --local user.email "actions@github.com"
          git config --local user.name  "EFC Bot"
          git add data/data.json
          if git diff --staged --quiet; then
            echo "✅ No data changes today."
          else
            git commit -m "🔄 Auto-update: $(TZ='Asia/Jakarta' date +'%Y-%m-%d %H:%M WIB')"
            git push
            echo "✅ data.json updated and pushed."
          fi
