#!/bin/bash
# ============================================================
# Seamfix Financial Intelligence Suite - One-Click Updater
# ============================================================
# HOW TO USE:
# 1. Drop your new weekly Excel file into the  data/  folder
#    (next to this file). Keep the existing filename pattern:
#    "Cash Report as at <date>.xlsx"
# 2. Double-click this file (UPDATE_DASHBOARD.command)
# 3. All 5 dashboards regenerate from data/ and open in your browser
#
# Note: this regenerates LOCAL preview HTML only. The hosted app
# (Streamlit Cloud) pulls live data from Google Drive/Sheets and is
# updated by pushing to GitHub — see CLAUDE.md.
# ============================================================

cd "$(dirname "$0")"
DATA_DIR="./data"

echo "=============================================="
echo "  Seamfix Financial Intelligence Suite"
echo "=============================================="
echo ""
echo "Reading data files from: $DATA_DIR"
ls -1 "$DATA_DIR"/*.xlsx 2>/dev/null
echo ""

echo "[1/5] Generating Cash Overview Dashboard..."
python3 generate_dashboard.py "$DATA_DIR"
echo ""

echo "[2/5] Generating Expense Analysis Dashboard..."
python3 generate_expense_dashboard.py "$DATA_DIR"
echo ""

echo "[3/5] Generating Budget vs Actual Dashboard..."
python3 generate_budget_dashboard.py "$DATA_DIR"
echo ""

echo "[4/5] Generating Revenue & Fundability Dashboard..."
python3 generate_revenue_dashboard.py "$DATA_DIR"
echo ""

echo "[5/5] Generating Pipeline Intelligence Dashboard..."
python3 generate_pipeline_dashboard.py "$DATA_DIR"
echo ""

echo "Done! Opening all dashboards..."
for f in dashboard budget_dashboard revenue_dashboard expense_dashboard pipeline_dashboard; do
    open "$DATA_DIR/$f.html" 2>/dev/null || xdg-open "$DATA_DIR/$f.html" 2>/dev/null || echo "Open $DATA_DIR/$f.html in your browser"
    sleep 1
done
