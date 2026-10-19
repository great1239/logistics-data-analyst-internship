# Logistics Data Analytics Internship - Enterprise Freight Optimization

**Candidate Name:** Pratham Sharma  
**Candidate Email:** `prathamsharma1239@gmail.com`  
**Enrollment ID:** `376280`  
**Role:** Logistics Data Analyst Intern  
**Platform:** [Yuva Intern](https://yuvaintern.com/my-internship/376280)  
**Internship Duration:** 4 Weeks (July 2026 – August 2026)  

---

## 📌 Project Overview
This repository contains the complete analytical codebase, technical documentation, visualization dashboards, and machine learning models developed during the **Logistics Data Analyst Internship** at **Yuva Intern**.

The project investigates freight transit operations across 5,000 multi-modal consignments spanning 8 key commercial hubs in India (**Delhi, Mumbai, Bengaluru, Chennai, Kolkata, Hyderabad, Pune, Ahmedabad**). The research addresses core operational challenges: transit delay volatility, warehouse cross-dock dwell times, environmental friction (weather and traffic), carrier reliability, and dynamic dispatch optimization.

---

## 📂 Repository Structure

```tree
├── data/
│   ├── raw_logistics_data.csv          # Raw ingested shipment records (5,000 rows)
│   └── cleaned_logistics_data.csv      # Imputed, cleaned & feature-engineered dataset
├── reports/
│   ├── Week_1_Strategic_Planning_and_Data_Exploration.docx
│   ├── Week_2_Data_Collection_Cleaning_and_Preprocessing.docx
│   ├── Week_3_Advanced_Data_Analysis_and_Visualization.docx
│   └── Week_4_Predictive_Modeling_and_Optimization.docx
├── scripts/
│   ├── generate_dataset.py             # Realistic logistics data synthesis
│   ├── week1_strategy.py               # Strategic scoping & Week 1 report compiler
│   ├── week2_preprocessing.py          # Cleaning, imputation & feature engineering
│   ├── week3_analysis_viz.py           # EDA, statistical analysis & figure generation
│   └── week4_predictive_modeling.py    # Regression & classification ML pipelines
├── visualizations/
│   ├── fig1_delay_distribution.png     # Transit delay distribution & SLA breach
│   ├── fig2_carrier_performance_boxplot.png # Carrier reliability boxplots
│   ├── fig3_correlation_heatmap.png    # Operational & financial correlation matrix
│   ├── fig4_weather_traffic_impact.png # Weather & traffic delay compounding
│   ├── fig5_top_bottleneck_corridors.png # Top 10 high-delay corridors
│   ├── fig6_feature_importance.png     # Gradient Boosting feature drivers
│   ├── fig7_roc_curve.png              # Receiver Operating Characteristic curves
│   └── fig8_actual_vs_predicted.png    # Model validation parity plot
├── requirements.txt                    # Python environment dependencies
└── README.md                           # Project documentation
```

---

## 🚀 Weekly Deliverables & Technical Synopsis

### Week 1: Strategic Planning and Data Exploration in Logistics
- **Objective:** Simulate enterprise logistics project scoping, KPI taxonomy, and analytical architecture.
- **Key Contributions:**
  - Defined 5 mission-critical KPIs: **On-Time Delivery Rate (OTD%)**, **Average Transit Delay (ATD)**, **Freight Cost per Ton-Km**, **Dock Dwell Duration**, and **Carrier Reliability Score**.
  - Formulated a 5-phase data science lifecycle spanning ingestion to prescriptive optimization.
  - Implemented baseline ingestion and schema audit pipelines in Python.
- **Deliverable:** `reports/Week_1_Strategic_Planning_and_Data_Exploration.docx`

### Week 2: Data Collection, Cleaning, and Preprocessing
- **Objective:** Build an end-to-end data preparation pipeline handling real-world sensor and logging noise.
- **Key Contributions:**
  - Imputed missing shipment weights using vehicle-stratified medians and warehouse dwell times using origin-hub medians.
  - Detected and capped extreme sensor telemetry spikes using **Interquartile Range (IQR)** boundaries ($Q3 + 3.0 \times IQR$).
  - Engineered 9 domain features: `Dispatch_Hour`, `Is_Rush_Hour`, `Is_Weekend`, `Route_Code`, `Cost_Per_Km`, `Cost_Per_Ton_Km`, and `Calculated_Speed_kmh`.
- **Deliverable:** `reports/Week_2_Data_Collection_Cleaning_and_Preprocessing.docx`

### Week 3: Advanced Data Analysis and Visualization
- **Objective:** Execute Exploratory Data Analysis (EDA) and construct diagnostic dashboards.
- **Key Contributions:**
  - Quantified right-skewed delay distribution with 38.4% consignments exceeding the 0.5-hour SLA threshold.
  - Benchmarked carrier partners, revealing BlueDart Logistics (88.2% OTD) as the most reliable, whereas Gati KWE exhibited highest variance ($IQR > 3.2$ hrs).
  - Evaluated weather-traffic interactions: Storm conditions with severe traffic inflated mean delays to 9.85 hours.
  - Isolated high-delay line-haul bottlenecks along Kolkata $\rightarrow$ Mumbai and Delhi $\rightarrow$ Bengaluru corridors.
- **Deliverable:** `reports/Week_3_Advanced_Data_Analysis_and_Visualization.docx` (with embedded high-resolution figures)

### Week 4: Predictive Modeling and Optimization in Logistics Systems
- **Objective:** Train supervised machine learning models for continuous delay forecasting and binary SLA breach risk classification.
- **Key Contributions:**
  - **Regression (Delay Hours):** Gradient Boosting Regressor achieved $R^2 = 0.81$, outperforming Ridge Regression and Random Forest Regressor.
  - **Classification (Late Delivery Risk):** Random Forest Classifier achieved **90.4% ROC-AUC** and **84.5% Accuracy**.
  - **Feature Drivers:** Warehouse dock dwell duration, storm weather, severe traffic, and carrier choice were identified as primary predictive drivers.
  - **Prescriptive Optimization:** Formulated algorithmic carrier reassignment rules and dynamic safety buffer calculations ($ETA_{dynamic} = ETA_{planned} + \hat{Delay} + 1.28 \sigma$).
- **Deliverable:** `reports/Week_4_Predictive_Modeling_and_Optimization.docx`

---

## 🛠️ Installation & Reproduction

```bash
# Clone the repository
git clone https://github.com/great1239/logistics-data-analyst-internship.git
cd logistics-data-analyst-internship

# Install dependencies
pip install -r requirements.txt

# Run the complete pipeline end-to-end
python scripts/generate_dataset.py
python scripts/week1_strategy.py
python scripts/week2_preprocessing.py
python scripts/week3_analysis_viz.py
python scripts/week4_predictive_modeling.py
```

---

## 📜 License & Acknowledgments
Developed as part of the **Yuva Intern** Logistics Data Analyst Internship Program (July 2026 – August 2026).
All deliverables comply with the formal submission requirements of Yuva Intern.
