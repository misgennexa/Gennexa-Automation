# 🚀 Nexgen Analytics Dashboard

A fully **offline** Excel-powered analytics dashboard built with Python + Streamlit.

---

## ▶️ How to Run from VS Code

### Method 1 — Press `F5` (Easiest)
1. Open this folder in VS Code
2. Press **F5** → select **"▶ Run Nexgen Dashboard"**
3. Open browser at **http://localhost:8501**

### Method 2 — Run Task (`Ctrl+Shift+B`)
1. Press **Ctrl+Shift+B**
2. It auto-runs **"🚀 Start Nexgen Dashboard"**
3. Open browser at **http://localhost:8501**

### Method 3 — VS Code Terminal
Open the integrated terminal (`Ctrl+`` ` ``) and run:
```powershell
.\.venv\Scripts\streamlit.exe run nexgen_app.py
```

### Method 4 — Double-click
Double-click **`launch_nexgen.bat`** — browser opens automatically.

---

## 📁 Project Structure

```
nexgen/
├── nexgen_app.py          ← Main app (all 7 pages)
├── launch_nexgen.bat      ← One-click Windows launcher
├── requirements.txt       ← Python dependencies
├── .venv/                 ← Python 3.12 virtual environment
├── .streamlit/
│   └── config.toml        ← Dark theme config
├── .vscode/
│   ├── launch.json        ← F5 run config
│   ├── tasks.json         ← Ctrl+Shift+B task
│   └── settings.json      ← Python interpreter + settings
├── templates/             ← Upload your Excel templates here
├── data/                  ← Daily saved dashboards (auto-created)
└── inbox/                 ← Drop files here for auto-import
```

---

## 🔄 Daily Workflow

1. **First time only** → Go to **Template Setup** page → Upload your Excel file with formulas
2. **Every day** → Go to **Upload Data** page → Drag in your raw CSV/Excel
3. Dashboard auto-builds with KPIs, charts, and trend lines
4. Export as **CSV / Excel / PDF** anytime

---

## 📦 Re-install dependencies (if needed)

```powershell
uv pip install --python .\.venv\Scripts\python.exe -r requirements.txt
```

---

## 🌐 Pages

| Page | Description |
|------|-------------|
| 📊 Dashboard | KPI cards, charts, trend lines |
| 📋 Template Setup | Upload Excel template with formulas |
| 📤 Upload Data | Daily data upload + column validation |
| 📅 History | Browse any previous day |
| 📈 Comparison | Day vs day / week over week |
| ✅ Data Validation | Blanks, duplicates, type errors |
| ⚙️ Settings | Thresholds, alerts, backup |
