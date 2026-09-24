# Leakage-Safe Explainable XGBoost Framework for Multiclass Maternal Health Risk Triage

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.20727538.svg)](https://doi.org/10.5281/zenodo.20727538)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Dataset:** 2023 Indonesian Health Survey (*Survei Kesehatan Indonesia - SKI 2023*) — Ministry of Health of the Republic of Indonesia ($N = 211,351$)  
> **Authors:** Attala Alif Ramadhani Tri Hida$^{1*}$, Wahyudi Setiawan$^{2}$  
> **Affiliation:** Department of Information Systems, Faculty of Engineering, Universitas Trunojoyo Madura  
> **Task:** Three-Tier Clinical Maternal Risk Stratification (Low Risk / High Risk / Very High Risk - KSPR Standard)  
> **Zenodo DOI:** [10.5281/zenodo.20727538](https://doi.org/10.5281/zenodo.20727538)

---

## 📌 Overview

This repository contains the complete, reproducible, and leakage-safe machine learning pipeline for classifying maternal health risk levels based on the official Indonesian **Poedji Rochjati Score Card (*Kartu Skor Poedji Rochjati*, KSPR)** triage standard. The framework is trained and validated on nationwide microdata from the 2023 Indonesian Health Survey (SKI 2023) covering 211,351 pregnant women across all 38 provinces.

### 🌟 Key Scientific & Methodological Contributions:

| Methodological Dimension | Technical Implementation & Clinical Safeguards |
|---|---|
| **Leakage-Safe Resampling** | Dynamic SMOTE oversampling is strictly encapsulated inside cross-validation training folds via `imblearn.pipeline.Pipeline`, preventing optimistic validation contamination. |
| **National Scale ($N=211,351$)** | Nationwide representative cohort across all 38 Indonesian provinces, resolving localized sample bias in prior literature. |
| **Three-Tier KSPR Alignment** | Directly operationalizes the clinical three-tier triage system (KRR, KRT, KRST) mandated by Indonesian primary healthcare protocols. |
| **Zero Fatal Under-Triage** | The proposed Tuned XGBoost model completely eliminates fatal false negatives from the Very High-Risk (KRST) category into Low Risk (0.0% fatal under-triage). |
| **Dual-Level Explainable AI** | **TreeSHAP** provides global population risk thresholds (maternal age $\ge 35$, miscarriage $\ge 1$, parity $\ge 4$) and individualized patient waterfall attributions. |
| **Statistical Significance** | 10-Fold Cross-Validation confirmed superiority via the **Wilcoxon Signed-Rank Test** ($W = 0.0000, p = 1.50 \times 10^{-4} < 0.001$). |
| **SDG Alignment** | Directly contributes to **SDG 3** (Maternal Mortality Reduction), **SDG 9** (Digital Health Innovation), and **SDG 10** (Geographic Healthcare Equity). |

---

## 📂 Repository Contents

```text
.
├── main.py       ← Single master pipeline, figure generator, and experimental report exporter
├── README.md     ← Documentation, methodological summary, and reproduction guide
└── .gitignore    ← Excludes raw survey CSVs, model weight caches, and generated media
```

> **Privacy & Reproducibility Notice:** Raw survey microdata from SKI 2023 is owned and managed by the [Health Data Service Portal of the Ministry of Health of Indonesia](https://layanandata.kemkes.go.id/) and cannot be redistributed publicly. Pre-trained model caches (`model_cache.pkl`) and generated figures are excluded per `.gitignore`.

---

## 🚀 Getting Started & Execution Guide

### 1. Installation of Dependencies

Ensure Python $\ge 3.9$ is installed, then install the required computational packages:

```bash
pip install pandas numpy matplotlib seaborn scikit-learn xgboost             imbalanced-learn shap scipy joblib pillow python-docx
```

### 2. Prepare the Dataset

Place your SKI 2023 dataset CSV file in the root repository directory and name it `dataset_ski_2023.csv`:

```text
.
├── dataset_ski_2023.csv
└── main.py
```

### 3. Run the Complete Pipeline

Execute the master pipeline:

```bash
python main.py
```

**Automated Pipeline Workflow:**
1. **Preprocessing & Leakage Prevention:** Loads microdata, filters validated pregnancies ($N = 211,351$), excludes non-predictive IDs and post-hoc surgical variables (`metode_persalinan_sesar`), and performs an 80/20 stratified split into training ($n = 169,080$) and hold-out test ($n = 42,271$) sets.
2. **Model Training & Hyperparameter Tuning:** Performs 5-fold stratified cross-validated randomized search for XGBoost and Random Forest, caching optimal models to `model_cache.pkl`.
3. **Statistical Validation:** Computes 10-fold CV metrics, per-class specificity, and Wilcoxon Signed-Rank hypothesis tests.
4. **Publication-Grade Figure Generation:** Renders high-resolution 300 DPI figures into:
   - `figures_combined/`: Figures 1–6 (side-by-side composite panels).
   - `figures_separated/`: Figures 1–10 (individual standalone plots for two-column journal templates).
5. **Report Export:** Automatically produces `laporan_hasil_eksperimen.docx` and `LAPORAN_HASIL_EKSPERIMEN.md`.

---

## 📊 Experimental Benchmark Summary

Evaluated on the independent nationwide hold-out test set ($n = 42,271$):

| Model Architecture | Accuracy (%) | F1-Macro | ROC-AUC (Macro OvR) | PR-AUC (Macro OvR) | Precision (KRST) (%) | Recall (KRST) (%) | F1-Score (KRST) |
|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| Decision Tree (C4.5 Baseline) | 93.78% | 0.8977 | 0.9273 | 0.9250 | 79.13% | 80.61% | 0.7986 |
| Random Forest (Default Baseline) | 94.77% | 0.9169 | 0.9931 | 0.9712 | 80.33% | 87.82% | 0.8391 |
| XGBoost (Default Baseline) | 95.10% | 0.9200 | 0.9944 | 0.9745 | 82.26% | 85.79% | 0.8399 |
| Random Forest (Tuned) | 94.80% | 0.9184 | 0.9934 | 0.9720 | 79.18% | **90.01%** | 0.8425 |
| **PROPOSED: Tuned XGBoost** | **95.16%** | **0.9207** | **0.9946** | **0.9754** | **82.25%** | 86.00% | **0.8408** |

### Clinical Safety & Statistical Significance:
* **Zero Fatal Under-Triage:** $0.00\%$ fatal false negatives (0 cases misclassified from KRST into Low Risk).
* **Wilcoxon Signed-Rank Test:** $W = 0.0000, p = 1.50 \times 10^{-4} < 0.001$ against Decision Tree and Random Forest baselines.

---

## ✍️ Citation & Zenodo Archival

If you utilize this codebase, methodology, or pipeline in your research, please cite:

```bibtex
@article{hida2026leakage,
  title={Leakage-Safe Explainable XGBoost Framework for Multiclass Maternal Health Risk Triage},
  author={Hida, Attala Alif Ramadhani Tri and Setiawan, Wahyudi},
  journal={Jurnal RESTI (Rekayasa Sistem dan Teknologi Informasi)},
  year={2026},
  doi={10.5281/zenodo.20727538},
  url={https://github.com/attaramadhani/laporan-crisp-dm}
}
```

* **GitHub Repository:** [https://github.com/attaramadhani/laporan-crisp-dm](https://github.com/attaramadhani/laporan-crisp-dm)  
* **Zenodo Archive DOI:** [https://doi.org/10.5281/zenodo.20727538](https://doi.org/10.5281/zenodo.20727538)
