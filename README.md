# Combined Mining Dashboard (Streamlit)

A Streamlit app that consolidates three tools into a single dashboard:

- Mining Data Processor — validates and summarizes daily mining report data.
- Production Dashboard — interactive charts/tables of production KPIs.
- Daily Report Update — generates/updates a Geology Work Plan workbook from a monthly Daily Report.

Main app entrypoint: `combined_app.py`.

## Features
- First-command page config safe from import-time conflicts.
- Robust module path handling for `v1.1.4/` and `v4/` modules.
- Uses an included template workbook in `v4/` (e.g. `Geology Daily Work Plan August2025.xlsx`).
- Works with uploaded Excel files and returns downloadable results.

## Project structure
```
.
├─ combined_app.py
├─ requirements.txt
├─ README.md
├─ jan_aug_data_with_bench_grades.xlsx               # sample data used by Production Dashboard
├─ v1.1.4/
│  ├─ mining_processor.py
│  ├─ app.py                                         # standalone legacy app (not imported by combined app)
│  ├─ config/
│  └─ processors/
└─ v4/
   ├─ update_daily_report_all_days.py
   ├─ August_2025_DAILY_REPORT.xlsx                  # example monthly report
   └─ Geology Daily Work Plan August2025.xlsx        # template used by the update page
```

## Local setup
Requirements: Python 3.10–3.11 recommended.

```bash
pip install -r requirements.txt
streamlit run combined_app.py
```

## Streamlit Community Cloud deployment
1. Push this folder to a public GitHub repository (see name ideas below).
2. Go to https://streamlit.io/cloud and create a new app.
3. Point to your repo/branch.
4. Set Main file path to `combined_app.py`.
5. Deploy.

Notes:
- The app reads the template from `v4/` by filename pattern `Geology Daily Work Plan*.xlsx`. Make sure the template file is committed.
- File writes occur in a temporary directory and downloads are served to the user (Cloud-compatible).
- If your Excel files are very large (>100 MB), consider Git LFS.

## Troubleshooting
- Page config errors: Ensure `st.set_page_config()` is the first Streamlit command in `combined_app.py` (already configured).
- Template not found: Confirm a file matching `Geology Daily Work Plan*.xlsx` exists in `v4/`.
- Import errors: The app auto-adds `v1.1.4/` and `v4/` to `sys.path`. Keep the directories alongside `combined_app.py`.

## License
Add your preferred license (e.g., MIT) if needed.
