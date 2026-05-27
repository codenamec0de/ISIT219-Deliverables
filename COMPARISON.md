# How this repo compares with Aiden's Orange prep

Quick note for the group, since Aiden has done some Orange setup work in his own `ISIT219_A2-main` folder and it's worth being clear about how the two pieces fit together.

Aiden's folder has an Orange workflow file plus a cleaned dataset. His README is honest about the state of it: "no formal analysis has been done yet, this is just the setup". So treating his folder as a prep workspace and this repo as the finished submission is the simplest way to think about it.

## What Aiden's folder covers

- Orange `.ows` workflow with the standard pipeline (File → Select Columns → Edit Domain → Preprocess → Test and Score)
- One classifier wired up (Logistic Regression) and one clusterer (K-Means at k = 2)
- Modified `.xls` with the X-label header row stripped so Orange parses it properly
- A few Box Plot and Scatter Plot widgets for visualisation
- Nice touch: he renamed the PAY_n category values to plain English ("paid early", "paid duly", "minimum payment", "1 month delay", and so on) which is more readable than the raw integers

## What this repo adds on top

Most of these are things the marking rubric explicitly asks for, so it's not really a competition with Aiden's work, it's the next stage.

- A finished 19-page report in PDF and DOCX form, ready to submit, written against the brief's four sections (Knowledge Acquisition, Knowledge Creation, Results and Discussion, Inconsistencies and Limitations) plus an intro and conclusion.
- Two classifiers (Logistic Regression and Random Forest), not one. The brief asks for at least two types per knowledge category, so a single classifier wouldn't tick that box.
- Two clusterers (K-Means and Ward hierarchical), with a contingency-table check between them showing 90.5% agreement on which client lands in which cluster. Same reason as above.
- Engineered features that summarise the six months of repayment history into single columns: number of months late, worst PAY status, mean PAY status, mean utilisation, payment-to-bill ratio, and a couple of others. These end up dominating the feature-importance ranking.
- 14 cited sources, all real and verifiable (Yeh and Lien for the canonical paper on this dataset, Lessmann for the benchmark, Brown and Mues, Khandani, and so on), in UOW Harvard format with a centred References page and hanging indent.
- A reproducible Python pipeline (`analysis.py` with seed 42) that you can re-run to get every number in the report. The metrics JSON is checked in so nothing is hand-typed.
- A proper inconsistencies and limitations section covering the undocumented EDUCATION / MARRIAGE codes, the PAY_n -2 and 0 ambiguity, what the class imbalance does to accuracy, the modest silhouette score, and how generalisable a 2005-Taiwan snapshot really is.

## Two things in the Orange workflow worth flagging before submission

1. **Class weighting is off in the Logistic Regression widget.** The target is only 22% positive, so without `class_weight = balanced` the LR will learn to predict "no default" most of the time. Its accuracy will look fine (around 78%) but it'll miss most actual defaulters. In this repo both models have balanced class weights and we report PR-AUC and recall instead of leaning on accuracy.
2. **PAY_n columns are converted to Categorical in Edit Domain.** That throws away the ordinal information — "3 months late" should rank above "2 months late" but Orange now treats those as unrelated categories. Tree models still pick it up because of one-hot encoding, but Logistic Regression loses the monotonicity entirely. In this repo PAY_n stays numeric.

Neither is a hard bug, they're just settings to flip if we ever want to revive the Orange version.

## What I borrowed from Aiden's setup

- The PAY_n English labels (paid early, paid duly, minimum payment, X month delay) — going to use these on the Figure 2 axis in the next revision because they read better than -2, -1, 0, 1, …
- The idea of renaming clusters to "higher risk of default" and "lower risk of default" instead of "Cluster 0" and "Cluster 1" — same reason, just clearer for a reader.

## Suggested approach for the group submission

- Use this repo as the submitted artefact (report.pdf goes in, code optional as a supplementary file).
- Keep Aiden's `.ows` workflow as a supplementary appendix file. It's good evidence that the team explored Orange too, and the contribution log can point at it.
- One sentence in the methods section of the report could acknowledge that the team prototyped in Orange before moving to Python. Easy to add if anyone wants to.

## If anything looks off

Same deal as Aiden's note — flag it in the group chat and we'll fix it before submission.
