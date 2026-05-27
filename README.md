# ISIT219 Assignment 2 — Credit Default Risk

Submitted version of our group report on the UCI Default of Credit Card Clients dataset, plus the Python pipeline that produced every number in it.

## What's in here

| File / folder | What it is |
|---|---|
| `report.pdf` | The submitted report (19 pages, UOW Harvard referencing). |
| `report.docx` | Same content as a Word doc, in case anyone needs to edit before final submission. |
| `analysis.py` | The pipeline. Loads the .xls, cleans it, fits two classifiers and two clusterers, writes all figures and metrics. Fixed random seed 42 so results are reproducible. |
| `build_report.py` | Builds `report.docx` from prose strings at the top of the file, then converting to PDF via LibreOffice. Edit the strings, re-run, done. |
| `data/` | The UCI .xls dataset (kept in the repo so the pipeline runs out of the box). |
| `figures/` | The 12 PNG figures that the report uses. Regenerated whenever `analysis.py` runs. |
| `results/` | `metrics.json` with every reported number, plus the threshold-sweep CSV and cluster-profile CSV. |
| `COMPARISON.md` | How this differs from the Orange workflow Aiden put together earlier. Worth reading before submission. |

## Group members

- Saif Abdulsattar (7978418)
- Aiden Gabriel Antonino (9903811)
- Faizan Ahmed Khan (7659076)
- Liam Mills (7971643)
- Huseyin Cagri Alaf (8096181)

Equal contribution: 20% each.

## How to run the pipeline

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install pandas xlrd scikit-learn matplotlib scipy numpy python-docx
python analysis.py         # regenerates figures/ and results/metrics.json
python build_report.py     # rebuilds report.docx (then convert to PDF however you like)
```

For the PDF conversion the easiest path is LibreOffice headless:

```bash
libreoffice --headless --convert-to pdf report.docx
```

## Headline results

- Random Forest: ROC-AUC 0.779, PR-AUC 0.557 on the held-out test set.
- Logistic Regression: ROC-AUC 0.746, PR-AUC 0.497.
- Two-cluster K-Means partition with 29.1% vs 14.2% observed default rates, confirmed by Ward hierarchical at 29.8% vs 14.1% (90.5% assignment agreement).
- Top predictors are all repayment-history summaries (PAY_0, max_pay_status, n_months_late). Demographics barely contribute.
- 
