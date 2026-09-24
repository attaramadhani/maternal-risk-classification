# Maternal Health Risk Classification (SKI 2023)

**Author:** Attala Alif Ramadhani Tri Hida  
**Affiliation:** Department of Information Systems, Faculty of Engineering, Universitas Trunojoyo Madura  
**Task:** Three-Tier Antenatal Risk Stratification (Low Risk / High Risk / Very High Risk - KSPR Standard)  
**Zenodo DOI:** https://doi.org/10.5281/zenodo.20727538  

---

## 1. Research Overview

This repository contains the machine learning pipeline for classifying maternal health risk levels based on the **Poedji Rochjati Score Card (KSPR)** triage standard, trained on the nationwide 2023 Indonesian Health Survey (*Survei Kesehatan Indonesia - SKI 2023*) dataset ($N = 211,351$).

To ensure scientific validity:
- Dynamic SMOTE oversampling is strictly encapsulated inside cross-validation training folds to prevent data leakage.
- Non-predictive administrative IDs and post-hoc cesarean delivery columns were removed from the predictor space (56 features).
- The pipeline benchmarks Decision Tree (C4.5), Random Forest, and Tuned XGBoost.
- Dual-Level TreeSHAP explains both global population risk boundaries and individual patient triage decisions.

---

## 2. Experimental Results

Evaluated on the independent hold-out test set ($n = 42,271$, 20% stratified partition):

### 2.1. Model Performance Benchmark

| Model Architecture | Accuracy (%) | F1-Macro | ROC-AUC (Macro OvR) | PR-AUC (Macro OvR) | Precision (KRST) (%) | Recall (KRST) (%) | F1-Score (KRST) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| Decision Tree (C4.5 Baseline) | 93.78% | 0.8977 | 0.9273 | 0.9250 | 79.13% | 80.61% | 0.7986 |
| Random Forest (Default Baseline) | 94.77% | 0.9169 | 0.9931 | 0.9712 | 80.33% | 87.82% | 0.8391 |
| XGBoost (Default Baseline) | 95.10% | 0.9200 | 0.9944 | 0.9745 | 82.26% | 85.79% | 0.8399 |
| Random Forest (Tuned) | 94.80% | 0.9184 | 0.9934 | 0.9720 | 79.18% | **90.01%** | 0.8425 |
| **PROPOSED: Tuned XGBoost** | **95.16%** | **0.9207** | **0.9946** | **0.9754** | **82.25%** | 86.00% | **0.8408** |

### 2.2. Clinical Safety & Under-Triage Analysis

In maternal emergency triage, misclassifying Very High-Risk (KRST) cases into the Low-Risk tier is a fatal error.
* **Tuned XGBoost Confusion Matrix on KRST ($n = 3,842$):**
  - Correctly Triaged to KRST: 3,304 (86.00%)
  - Triaged to High Risk: 538 (14.00%)
  - **Misclassified to Low Risk: 0 cases (0.00% fatal under-triage)**

### 2.3. Statistical Significance Testing

10-Fold Cross-Validation confirmed performance stability across all folds:
* **Tuned XGBoost:** Accuracy $0.9516 \pm 0.0011$ | F1-Macro $0.9207 \pm 0.0013$
* **Tuned Random Forest:** Accuracy $0.9480 \pm 0.0012$ | F1-Macro $0.9184 \pm 0.0015$
* **Wilcoxon Signed-Rank Test:** $W = 0.0000$, $p = 1.50 \times 10^{-4}$ ($p < 0.001$). The superiority of Tuned XGBoost is statistically significant.

### 2.4. Explainable AI (SHAP) Insights

* **Maternal Age:** Risk attribution turns sharply positive at $\ge 35$ years ($\phi_i > +0.5$).
* **Miscarriage History:** Even a single prior miscarriage ($\ge 1$) causes an immediate surge in risk ($\phi_i = +1.18$).
* **Parity:** Grand multiparity ($\ge 4$ births) triggers heightened risk attribution ($\phi_i > +0.8$).

---

## 3. How to Run

### Requirements
```bash
pip install pandas numpy matplotlib seaborn scikit-learn xgboost imbalanced-learn shap scipy joblib pillow python-docx
```

### Execution
Place `dataset_ski_2023.csv` in the project directory, then run:
```bash
python main.py
```

The script will automatically train/load the models, perform statistical validation, generate 300 DPI figures (`figures_combined/` and `figures_separated/`), and export the full evaluation report (`experimental_results_report.docx`).
