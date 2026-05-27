"""ISIT219 Assignment 2 — full analysis pipeline.

Loads UCI Default of Credit Card Clients, prepares features, trains two
classifiers (Logistic Regression + Random Forest), runs two clusterers
(K-Means + Ward Hierarchical), writes every figure used in the report
to ./figures and every numeric result to ./results/metrics.json.

Run:  /tmp/isit_env/bin/python analysis.py
"""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.cluster.hierarchy import dendrogram, fcluster, linkage
from sklearn.cluster import KMeans
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    average_precision_score,
    classification_report,
    confusion_matrix,
    f1_score,
    precision_recall_curve,
    roc_auc_score,
    roc_curve,
    silhouette_score,
)
from sklearn.model_selection import StratifiedKFold, cross_val_score, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

SEED = 42
BUILD = Path(__file__).parent
FIG = BUILD / "figures"
RES = BUILD / "results"
XLS = BUILD / "data" / "default of credit card clients.xls"

plt.rcParams.update({"figure.dpi": 130, "savefig.dpi": 130, "font.size": 10})


def load_and_clean() -> pd.DataFrame:
    df = pd.read_excel(XLS, header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={"default payment next month": "default"})
    df = df.drop(columns=["ID"])

    df["EDUCATION"] = df["EDUCATION"].replace({0: 4, 5: 4, 6: 4})
    df["MARRIAGE"] = df["MARRIAGE"].replace({0: 3})

    pay_cols = ["PAY_0", "PAY_2", "PAY_3", "PAY_4", "PAY_5", "PAY_6"]
    bill_cols = [f"BILL_AMT{i}" for i in range(1, 7)]
    payamt_cols = [f"PAY_AMT{i}" for i in range(1, 7)]

    df["mean_pay_status"] = df[pay_cols].mean(axis=1)
    df["max_pay_status"] = df[pay_cols].max(axis=1)
    df["n_months_late"] = (df[pay_cols] >= 1).sum(axis=1)
    df["any_severe_delinq"] = (df[pay_cols] >= 3).any(axis=1).astype(int)

    util = df[bill_cols].div(df["LIMIT_BAL"].replace(0, np.nan), axis=0).clip(lower=-1, upper=5)
    df["mean_utilisation"] = util.mean(axis=1).fillna(0)
    df["max_utilisation"] = util.max(axis=1).fillna(0)

    df["total_bill_6m"] = df[bill_cols].sum(axis=1)
    df["total_paid_6m"] = df[payamt_cols].sum(axis=1)
    df["payment_to_bill"] = np.where(
        df["total_bill_6m"].abs() > 1.0,
        df["total_paid_6m"] / df["total_bill_6m"].replace(0, np.nan),
        0.0,
    )
    df["payment_to_bill"] = df["payment_to_bill"].replace([np.inf, -np.inf], 0).fillna(0).clip(-2, 5)
    df["bill_trend"] = df["BILL_AMT1"] - df["BILL_AMT6"]
    return df


def eda_figures(df: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(4.5, 3.2))
    counts = df["default"].value_counts().sort_index()
    bars = ax.bar(["Non-default", "Default"], counts.values, color=["#4c78a8", "#e45756"])
    for b, v in zip(bars, counts.values):
        ax.text(b.get_x() + b.get_width() / 2, v + 200, f"{v:,}\n({v / len(df):.1%})", ha="center", fontsize=9)
    ax.set_ylabel("Clients")
    ax.set_title("Class distribution (target)")
    ax.set_ylim(0, counts.max() * 1.18)
    fig.tight_layout()
    fig.savefig(FIG / "01_class_balance.png")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(6, 3.6))
    grp = df.groupby("PAY_0")["default"].agg(["mean", "count"]).reset_index()
    ax.bar(grp["PAY_0"].astype(str), grp["mean"], color="#5b8cb6")
    for x, (m, n) in enumerate(zip(grp["mean"], grp["count"])):
        ax.text(x, m + 0.01, f"n={n:,}", ha="center", fontsize=7, rotation=0)
    ax.set_xlabel("PAY_0 (September 2005 repayment status)")
    ax.set_ylabel("Default rate")
    ax.set_title("Default rate by most recent month's repayment status")
    ax.set_ylim(0, 1)
    fig.tight_layout()
    fig.savefig(FIG / "02_default_by_pay0.png")
    plt.close(fig)

    edu_label = {1: "Graduate", 2: "University", 3: "High school", 4: "Other"}
    grp = df.groupby("EDUCATION")["default"].agg(["mean", "count"]).reset_index()
    grp["label"] = grp["EDUCATION"].map(edu_label)
    fig, ax = plt.subplots(figsize=(5, 3.4))
    ax.bar(grp["label"], grp["mean"], color="#8c6bb1")
    for x, (m, n) in enumerate(zip(grp["mean"], grp["count"])):
        ax.text(x, m + 0.005, f"n={n:,}\n{m:.1%}", ha="center", fontsize=8)
    ax.set_ylabel("Default rate")
    ax.set_title("Default rate by education level")
    ax.set_ylim(0, max(grp["mean"]) * 1.4)
    fig.tight_layout()
    fig.savefig(FIG / "03_default_by_education.png")
    plt.close(fig)

    bins = [20, 30, 40, 50, 60, 80]
    labels = ["21-30", "31-40", "41-50", "51-60", "61-79"]
    df["age_band"] = pd.cut(df["AGE"], bins=bins, labels=labels, right=True, include_lowest=True)
    grp = df.groupby("age_band", observed=True)["default"].agg(["mean", "count"]).reset_index()
    fig, ax = plt.subplots(figsize=(5, 3.4))
    ax.bar(grp["age_band"].astype(str), grp["mean"], color="#6aa56e")
    for x, (m, n) in enumerate(zip(grp["mean"], grp["count"])):
        ax.text(x, m + 0.005, f"n={n:,}\n{m:.1%}", ha="center", fontsize=8)
    ax.set_ylabel("Default rate")
    ax.set_title("Default rate by age band")
    ax.set_ylim(0, max(grp["mean"]) * 1.4)
    fig.tight_layout()
    fig.savefig(FIG / "04_default_by_age.png")
    plt.close(fig)
    df.drop(columns=["age_band"], inplace=True)

    keep = ["LIMIT_BAL", "AGE", "PAY_0", "PAY_2", "PAY_3", "mean_pay_status",
            "max_pay_status", "n_months_late", "mean_utilisation",
            "payment_to_bill", "default"]
    corr = df[keep].corr()
    fig, ax = plt.subplots(figsize=(6.5, 5.5))
    im = ax.imshow(corr, cmap="RdBu_r", vmin=-1, vmax=1)
    ax.set_xticks(range(len(keep)))
    ax.set_xticklabels(keep, rotation=45, ha="right", fontsize=8)
    ax.set_yticks(range(len(keep)))
    ax.set_yticklabels(keep, fontsize=8)
    for i in range(len(keep)):
        for j in range(len(keep)):
            ax.text(j, i, f"{corr.iloc[i, j]:.2f}", ha="center", va="center",
                    color="white" if abs(corr.iloc[i, j]) > 0.5 else "black", fontsize=7)
    ax.set_title("Correlation among selected features (Pearson)")
    fig.colorbar(im, ax=ax, shrink=0.7)
    fig.tight_layout()
    fig.savefig(FIG / "05_correlation_heatmap.png")
    plt.close(fig)


def build_pipelines(num_cols, cat_cols):
    pre = ColumnTransformer(
        [("num", StandardScaler(), num_cols),
         ("cat", OneHotEncoder(drop="first", handle_unknown="ignore"), cat_cols)]
    )
    lr = Pipeline([
        ("pre", pre),
        ("clf", LogisticRegression(max_iter=2000, class_weight="balanced",
                                    solver="liblinear", C=1.0, random_state=SEED)),
    ])
    rf = Pipeline([
        ("pre", pre),
        ("clf", RandomForestClassifier(n_estimators=400, max_depth=12,
                                        min_samples_leaf=20, n_jobs=-1,
                                        class_weight="balanced_subsample",
                                        random_state=SEED)),
    ])
    return lr, rf


def evaluate(name, model, X_te, y_te, results):
    proba = model.predict_proba(X_te)[:, 1]
    pred = (proba >= 0.5).astype(int)
    auc = roc_auc_score(y_te, proba)
    ap = average_precision_score(y_te, proba)
    f1 = f1_score(y_te, pred)
    cm = confusion_matrix(y_te, pred)
    report = classification_report(y_te, pred, output_dict=True, digits=3)

    fpr, tpr, _ = roc_curve(y_te, proba)
    prec, rec, _ = precision_recall_curve(y_te, proba)

    results["classification"][name] = {
        "roc_auc": round(float(auc), 4),
        "pr_auc": round(float(ap), 4),
        "f1_at_0.5": round(float(f1), 4),
        "accuracy": round(float(report["accuracy"]), 4),
        "precision_default": round(float(report["1"]["precision"]), 4),
        "recall_default": round(float(report["1"]["recall"]), 4),
        "f1_default": round(float(report["1"]["f1-score"]), 4),
        "confusion_matrix": cm.tolist(),
    }
    return {"name": name, "fpr": fpr, "tpr": tpr, "auc": auc,
            "prec": prec, "rec": rec, "ap": ap, "cm": cm, "proba": proba}


def classification_block(df: pd.DataFrame, results: dict) -> dict:
    target = "default"
    cat_cols = ["SEX", "EDUCATION", "MARRIAGE"]
    base_num = ["LIMIT_BAL", "AGE",
                "PAY_0", "PAY_2", "PAY_3", "PAY_4", "PAY_5", "PAY_6",
                "BILL_AMT1", "BILL_AMT2", "BILL_AMT3", "BILL_AMT4", "BILL_AMT5", "BILL_AMT6",
                "PAY_AMT1", "PAY_AMT2", "PAY_AMT3", "PAY_AMT4", "PAY_AMT5", "PAY_AMT6",
                "mean_pay_status", "max_pay_status", "n_months_late", "any_severe_delinq",
                "mean_utilisation", "max_utilisation", "total_bill_6m", "total_paid_6m",
                "payment_to_bill", "bill_trend"]

    X = df[base_num + cat_cols]
    y = df[target]
    X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.2,
                                              stratify=y, random_state=SEED)

    lr, rf = build_pipelines(base_num, cat_cols)

    skf = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
    cv_lr = cross_val_score(lr, X_tr, y_tr, cv=skf, scoring="roc_auc", n_jobs=-1)
    cv_rf = cross_val_score(rf, X_tr, y_tr, cv=skf, scoring="roc_auc", n_jobs=-1)
    results["classification"]["cv_roc_auc"] = {
        "logistic_regression_mean": round(float(cv_lr.mean()), 4),
        "logistic_regression_std": round(float(cv_lr.std()), 4),
        "random_forest_mean": round(float(cv_rf.mean()), 4),
        "random_forest_std": round(float(cv_rf.std()), 4),
    }

    lr.fit(X_tr, y_tr)
    rf.fit(X_tr, y_tr)

    lr_out = evaluate("logistic_regression", lr, X_te, y_te, results)
    rf_out = evaluate("random_forest", rf, X_te, y_te, results)

    fig, ax = plt.subplots(figsize=(5, 4))
    ax.plot(lr_out["fpr"], lr_out["tpr"], label=f"LR (AUC={lr_out['auc']:.3f})", color="#4c78a8")
    ax.plot(rf_out["fpr"], rf_out["tpr"], label=f"RF (AUC={rf_out['auc']:.3f})", color="#e45756")
    ax.plot([0, 1], [0, 1], "k--", lw=0.8)
    ax.set_xlabel("False positive rate")
    ax.set_ylabel("True positive rate")
    ax.set_title("ROC curves on held-out test set")
    ax.legend(loc="lower right")
    fig.tight_layout()
    fig.savefig(FIG / "06_roc_curves.png")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(5, 4))
    ax.plot(lr_out["rec"], lr_out["prec"], label=f"LR (AP={lr_out['ap']:.3f})", color="#4c78a8")
    ax.plot(rf_out["rec"], rf_out["prec"], label=f"RF (AP={rf_out['ap']:.3f})", color="#e45756")
    baseline = float(y_te.mean())
    ax.axhline(baseline, color="grey", lw=0.7, ls="--", label=f"Baseline = {baseline:.3f}")
    ax.set_xlabel("Recall (defaults)")
    ax.set_ylabel("Precision (defaults)")
    ax.set_title("Precision-recall curves on held-out test set")
    ax.legend(loc="upper right")
    fig.tight_layout()
    fig.savefig(FIG / "07_pr_curves.png")
    plt.close(fig)

    cm = rf_out["cm"]
    fig, ax = plt.subplots(figsize=(4.2, 3.6))
    im = ax.imshow(cm, cmap="Blues")
    ax.set_xticks([0, 1]); ax.set_xticklabels(["Pred 0", "Pred 1"])
    ax.set_yticks([0, 1]); ax.set_yticklabels(["True 0", "True 1"])
    for i in range(2):
        for j in range(2):
            ax.text(j, i, f"{cm[i, j]:,}", ha="center", va="center",
                    color="white" if cm[i, j] > cm.max() / 2 else "black", fontsize=11)
    ax.set_title("Random Forest confusion matrix (test, t=0.5)")
    fig.colorbar(im, ax=ax, shrink=0.7)
    fig.tight_layout()
    fig.savefig(FIG / "08_confusion_matrix_rf.png")
    plt.close(fig)

    rf_model = rf.named_steps["clf"]
    feat_names = list(base_num) + list(
        rf.named_steps["pre"].named_transformers_["cat"].get_feature_names_out(cat_cols)
    )
    importances = pd.Series(rf_model.feature_importances_, index=feat_names).sort_values(ascending=False)
    top = importances.head(15)[::-1]
    fig, ax = plt.subplots(figsize=(6, 5))
    ax.barh(top.index, top.values, color="#3a7d44")
    ax.set_xlabel("Feature importance (mean decrease in impurity)")
    ax.set_title("Random Forest — top 15 features")
    fig.tight_layout()
    fig.savefig(FIG / "09_feature_importance_rf.png")
    plt.close(fig)
    results["classification"]["top_feature_importances"] = {
        k: round(float(v), 4) for k, v in importances.head(15).items()
    }

    lr_model = lr.named_steps["clf"]
    coefs = pd.Series(lr_model.coef_[0], index=feat_names).sort_values()
    results["classification"]["lr_top_positive_coefs"] = {
        k: round(float(v), 4) for k, v in coefs.tail(8)[::-1].items()
    }
    results["classification"]["lr_top_negative_coefs"] = {
        k: round(float(v), 4) for k, v in coefs.head(8).items()
    }

    thresholds = np.linspace(0.05, 0.95, 19)
    rows = []
    for t in thresholds:
        p = (rf_out["proba"] >= t).astype(int)
        cm_t = confusion_matrix(y_te, p)
        tn, fp, fn, tp = cm_t.ravel()
        prec = tp / max(tp + fp, 1)
        rec = tp / max(tp + fn, 1)
        f = 2 * prec * rec / max(prec + rec, 1e-9)
        rows.append((float(t), float(prec), float(rec), float(f)))
    rows_df = pd.DataFrame(rows, columns=["threshold", "precision", "recall", "f1"])
    rows_df.to_csv(RES / "rf_threshold_sweep.csv", index=False)
    best = rows_df.loc[rows_df["f1"].idxmax()]
    results["classification"]["rf_best_f1_threshold"] = {
        "threshold": round(float(best["threshold"]), 3),
        "precision": round(float(best["precision"]), 4),
        "recall": round(float(best["recall"]), 4),
        "f1": round(float(best["f1"]), 4),
    }
    return {"rf": rf, "lr": lr, "X_te": X_te, "y_te": y_te}


def clustering_block(df: pd.DataFrame, results: dict) -> None:
    cluster_features = ["LIMIT_BAL", "AGE", "mean_pay_status", "max_pay_status",
                         "n_months_late", "mean_utilisation", "payment_to_bill",
                         "total_bill_6m", "total_paid_6m"]
    X = df[cluster_features].values
    scaler = StandardScaler()
    Xs = scaler.fit_transform(X)

    rng = np.random.default_rng(SEED)
    sample_idx = rng.choice(len(Xs), size=min(8000, len(Xs)), replace=False)
    Xs_sample = Xs[sample_idx]

    inertias, silhouettes = [], []
    ks = list(range(2, 9))
    for k in ks:
        km = KMeans(n_clusters=k, n_init=10, random_state=SEED)
        labels = km.fit_predict(Xs_sample)
        inertias.append(float(km.inertia_))
        silhouettes.append(float(silhouette_score(Xs_sample, labels, sample_size=4000, random_state=SEED)))

    fig, axes = plt.subplots(1, 2, figsize=(9, 3.4))
    axes[0].plot(ks, inertias, marker="o", color="#4c78a8")
    axes[0].set_xlabel("k"); axes[0].set_ylabel("Inertia")
    axes[0].set_title("K-Means elbow plot")
    axes[1].plot(ks, silhouettes, marker="o", color="#e45756")
    axes[1].set_xlabel("k"); axes[1].set_ylabel("Silhouette score")
    axes[1].set_title("Silhouette by k")
    fig.tight_layout()
    fig.savefig(FIG / "10_kmeans_selection.png")
    plt.close(fig)

    best_k = ks[int(np.argmax(silhouettes))]
    results["clustering"]["kmeans_selection"] = {
        "ks": ks,
        "inertias": [round(v, 1) for v in inertias],
        "silhouettes": [round(v, 4) for v in silhouettes],
        "selected_k": best_k,
    }

    km = KMeans(n_clusters=best_k, n_init=20, random_state=SEED)
    df["kmeans_cluster"] = km.fit_predict(Xs)
    profile_cols = cluster_features + ["default"]
    profile = df.groupby("kmeans_cluster")[profile_cols].mean().round(2)
    profile["size"] = df["kmeans_cluster"].value_counts().sort_index().values
    profile["default_rate"] = df.groupby("kmeans_cluster")["default"].mean().round(4).values
    profile.to_csv(RES / "kmeans_profiles.csv")
    results["clustering"]["kmeans_profiles"] = profile.reset_index().to_dict(orient="records")

    fig, ax = plt.subplots(figsize=(5, 3.4))
    rates = df.groupby("kmeans_cluster")["default"].mean().sort_index()
    sizes = df["kmeans_cluster"].value_counts().sort_index()
    bars = ax.bar([f"C{c}" for c in rates.index], rates.values, color="#5b8cb6")
    for b, r, s in zip(bars, rates.values, sizes.values):
        ax.text(b.get_x() + b.get_width() / 2, r + 0.005,
                f"{r:.1%}\nn={s:,}", ha="center", fontsize=8)
    ax.set_ylabel("Observed default rate")
    ax.set_ylim(0, max(rates.values) * 1.4)
    ax.set_title(f"K-Means (k={best_k}) — default rate per cluster")
    fig.tight_layout()
    fig.savefig(FIG / "11_kmeans_default_rate.png")
    plt.close(fig)

    rng = np.random.default_rng(SEED + 1)
    hier_idx = rng.choice(len(Xs), size=4000, replace=False)
    Z = linkage(Xs[hier_idx], method="ward")
    fig, ax = plt.subplots(figsize=(8, 3.6))
    dendrogram(Z, no_labels=True, color_threshold=0, ax=ax,
               above_threshold_color="#4c78a8")
    ax.set_title("Hierarchical clustering (Ward linkage, n=4 000 sample)")
    ax.set_xlabel("Sampled clients")
    ax.set_ylabel("Distance")
    fig.tight_layout()
    fig.savefig(FIG / "12_dendrogram.png")
    plt.close(fig)

    hier_k = best_k
    hier_labels = fcluster(Z, t=hier_k, criterion="maxclust")
    hier_default = (df["default"].iloc[hier_idx]
                    .groupby(hier_labels).mean().round(4).to_dict())
    hier_sizes = pd.Series(hier_labels).value_counts().sort_index().to_dict()
    results["clustering"]["hierarchical"] = {
        "k": hier_k,
        "sample_size": len(hier_idx),
        "default_rate_by_cluster": {int(k): float(v) for k, v in hier_default.items()},
        "cluster_sizes": {int(k): int(v) for k, v in hier_sizes.items()},
    }

    contingency = (pd.crosstab(df["kmeans_cluster"].iloc[hier_idx], hier_labels)
                   .to_dict())
    results["clustering"]["hier_vs_kmeans_contingency"] = {
        str(k): {int(c): int(n) for c, n in v.items()} for k, v in contingency.items()
    }


def main() -> None:
    print("Loading dataset…")
    df = load_and_clean()
    results = {"data": {"n_rows": int(len(df)),
                         "n_cols": int(df.shape[1]),
                         "default_rate": round(float(df["default"].mean()), 4),
                         "edu_after_recode": df["EDUCATION"].value_counts().to_dict(),
                         "marriage_after_recode": df["MARRIAGE"].value_counts().to_dict()},
                "classification": {}, "clustering": {}}

    print("EDA figures…")
    eda_figures(df)

    print("Classification…")
    classification_block(df, results)

    print("Clustering…")
    clustering_block(df, results)

    out = RES / "metrics.json"
    out.write_text(json.dumps(results, indent=2, default=str))
    print(f"Done. Metrics → {out}")


if __name__ == "__main__":
    main()
