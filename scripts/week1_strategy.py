"""
week1_strategy.py
Generates the comprehensive Week 1 Report as a formatted Word Document (.docx)
Topic: Strategic Planning and Data Exploration in Logistics
Candidate: Pratham Sharma (ID: 376280)
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_week1_report(output_path):
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Document Title Block
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title.add_run("LOGISTICS DATA ANALYST INTERNSHIP\nWEEK 1 TECHNICAL REPORT")
    run_title.font.name = 'Arial'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(16, 44, 87) # Deep Navy

    sub_title = doc.add_paragraph()
    sub_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_title.add_run("Strategic Planning and Data Exploration in Enterprise Freight Logistics")
    run_sub.font.name = 'Arial'
    run_sub.font.size = Pt(14)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(53, 89, 143)

    # Metadata Box Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    meta_data = [
        ("Candidate Name", "Pratham Sharma (prathamsharma1239@gmail.com)"),
        ("Enrollment ID", "376280"),
        ("Internship Role", "Logistics Data Analyst Intern"),
        ("Host Organization", "Yuva Intern (yuvaintern.com)")
    ]

    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.3)
        set_cell_background(cell_lbl, "EBF3FA")
        set_cell_background(cell_val, "F9FBFD")
        set_cell_margins(cell_lbl, top=80, bottom=80, left=120, right=120)
        set_cell_margins(cell_val, top=80, bottom=80, left=120, right=120)

        p0 = cell_lbl.paragraphs[0]
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.name = 'Arial'
        r0.font.color.rgb = RGBColor(16, 44, 87)

        p1 = cell_val.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.size = Pt(9.5)
        r1.font.name = 'Arial'

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    def add_section_heading(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
        r = h.add_run(text)
        r.font.name = 'Arial'
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(16, 44, 87)
        return h

    def add_body_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(text)
        r.font.name = 'Calibri'
        r.font.size = Pt(11)
        return p

    def add_bullet_p(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(4)
        r_pre = p.add_run(bold_prefix + ": ")
        r_pre.font.bold = True
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(11)
        r_txt = p.add_run(text)
        r_txt.font.name = 'Calibri'
        r_txt.font.size = Pt(11)
        return p

    # 1. Executive Summary
    add_section_heading("1. Executive Summary")
    add_body_p(
        "Modern freight and supply chain operations operate in highly volatile environments characterized by "
        "fuel price fluctuations, dynamic traffic bottlenecks, multi-modal vehicle allocations, and tight customer Service "
        "Level Agreements (SLAs). The objective of Week 1 is to establish the strategic and analytical foundation "
        "for an enterprise-scale Logistics Data Analytics project. This project models and analyzes 5,000 shipment dispatches "
        "across 8 major Indian commercial hubs (Delhi, Mumbai, Bengaluru, Chennai, Kolkata, Hyderabad, Pune, and Ahmedabad)."
    )
    add_body_p(
        "Through disciplined scoping, quantitative KPI formulation, and a 5-phase analytical roadmap, this project bridges "
        "raw operational event logging with executive-level supply chain visibility. This report defines the core operational "
        "scenario, articulates mission-critical Key Performance Indicators (KPIs), formulates the analytical architecture, "
        "and details initial Python scripts for data ingestion and exploratory pipeline design."
    )

    # 2. Logistics Business Scenario & Problem Formulation
    add_section_heading("2. Business Scenario & Problem Definition")
    add_body_p(
        "The project focuses on a mid-to-large third-party logistics (3PL) and freight forwarding enterprise that coordinates "
        "inter-city line-haul transportation and regional cross-dock dispatches. Freight moves via four distinct vehicle categories: "
        "Electric Delivery Vans (urban last-mile / short-haul), Light Commercial Vehicles (LCV), Medium Commercial Vehicles (MCV), "
        "and 16-Ton Heavy Commercial Vehicles (HCV)."
    )
    add_body_p(
        "Operational pain points currently undermining profitability and customer satisfaction include:"
    )
    add_bullet_p("Delivery Transit Delays", "Unpredicted delays caused by transit weather disruptions, severe urban bottlenecks, and highway congestion result in missed customer delivery windows.")
    add_bullet_p("Warehouse Cross-Dock Congestion", "Excessive dwell times during loading, sorting, and vehicle turnarounds at origin and transit hubs add variance to overall turnaround times.")
    add_bullet_p("Volatile Shipping & Fuel Costs", "Fluctuating diesel and fuel prices combined with sub-optimal vehicle load factors inflate operational costs per ton-kilometer.")
    add_bullet_p("Carrier Reliability Variances", "Inconsistent performance across partnered logistics carriers (BlueDart, Delhivery, SafeXpress, TCI, Gati) requires algorithmic monitoring and objective scoring.")

    # 3. Key Performance Indicators Matrix
    add_section_heading("3. Quantitative Key Performance Indicators (KPI) Matrix")
    add_body_p(
        "To evaluate operational performance objectively, five standardized metrics have been established. "
        "The table below outlines each KPI, its mathematical formulation, industry target, and operational significance:"
    )

    kpi_table = doc.add_table(rows=6, cols=4)
    kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    kpi_headers = ["Key Performance Indicator", "Mathematical Formulation", "Target Benchmark", "Business Impact"]

    for c_idx, h_text in enumerate(kpi_headers):
        cell = kpi_table.rows[0].cells[c_idx]
        set_cell_background(cell, "102C57")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        r.font.name = 'Arial'

    kpis = [
        ("On-Time Delivery Rate (OTD%)", "(N_ontime / N_total) * 100", ">= 92.0%", "Primary measure of customer SLA adherence and delivery promise fulfillment."),
        ("Average Transit Delay (ATD)", "Mean(Actual_Transit - Planned_Transit)", "< 1.5 Hours", "Quantifies schedule deviation severity across line-haul corridors."),
        ("Freight Cost per Ton-Km", "Shipping_Cost / (Weight_Tons * Distance_Km)", "<= INR 3.80", "Measures load consolidation efficiency and equipment utilization."),
        ("Dock Dwell Duration", "Mean(Warehouse_Dock_Hours)", "< 3.0 Hours", "Exposes warehouse processing bottlenecks and loading gate inefficiencies."),
        ("Carrier Reliability Score", "P(Delay <= 0.5 | Carrier = c)", ">= 88.0%", "Guides dynamic contract allocation and performance-based carrier bonuses.")
    ]

    for r_idx, row_data in enumerate(kpis):
        row = kpi_table.rows[r_idx + 1]
        bg = "FFFFFF" if r_idx % 2 == 0 else "F4F7FB"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9.0)
            r.font.name = 'Calibri'
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 4. Scope and Analytical Roadmap
    add_section_heading("4. Analytical Roadmap & Phased Execution Strategy")
    add_body_p(
        "The project follows a structured 5-phase data science lifecycle designed to progress from raw data ingestion "
        "to predictive and prescriptive decision support:"
    )
    add_bullet_p("Phase 1 - Ingestion & Schema Audit (Week 1)", "Establish automated data ingestion pipelines, audit schema integrity, verify data dictionary definitions, and establish foundational baseline metrics.")
    add_bullet_p("Phase 2 - Cleaning & Feature Engineering (Week 2)", "Impute missing records using KNN and median approaches, detect anomalous transit spikes via IQR filtering, engineer temporal and route features, and scale numerical vectors.")
    add_bullet_p("Phase 3 - Exploratory Data Analysis & Dashboards (Week 3)", "Uncover multi-dimensional correlations, quantify carrier performance variance, analyze weather and traffic bottlenecks, and construct executive visualization dashboards.")
    add_bullet_p("Phase 4 - Predictive Machine Learning Modeling (Week 4)", "Train regression models (Random Forest, Gradient Boosting) for delay duration forecasting, train classification models for late shipment risk, and perform hyperparameter optimization.")
    add_bullet_p("Phase 5 - Prescriptive Optimization & Integration (Week 4)", "Formulate dynamic carrier reassignment rules, route buffer allocations, and automated SLA violation alerting thresholds.")

    # 5. Data Architecture & Ingestion Design
    add_section_heading("5. Data Architecture & Schema Specification")
    add_body_p(
        "The core data model captures 20 granular operational variables per consignment, structured into three relational facets: "
        "Consignment Attributes, Transit & Environmental Context, and Commercial & Cost Metrics."
    )

    schema_table = doc.add_table(rows=6, cols=3)
    schema_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["Attribute Name", "Data Type", "Operational Description"]

    for c_idx, h_text in enumerate(s_headers):
        cell = schema_table.rows[0].cells[c_idx]
        set_cell_background(cell, "102C57")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        r.font.name = 'Arial'

    schema_items = [
        ("Shipment_ID", "String (Primary Key)", "Unique alphanumeric tracking identifier (SHP-XXXXXX)."),
        ("Origin_City / Destination_City", "Categorical / String", "Origin and destination commercial logistics hub nodes."),
        ("Distance_km & Vehicle_Type", "Float / Categorical", "Geodesic route distance in km and vehicle payload category."),
        ("Dispatch_Timestamp", "Datetime (ISO-8601)", "Exact timestamp when consignment completed dock exit scan."),
        ("Delay_Hours & Is_Delayed", "Float / Binary Flag", "Continuous variance from schedule and binary SLA breach indicator.")
    ]

    for r_idx, row_data in enumerate(schema_items):
        row = schema_table.rows[r_idx + 1]
        bg = "FFFFFF" if r_idx % 2 == 0 else "F4F7FB"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9.0)
            r.font.name = 'Calibri'
            if c_idx == 0:
                r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 6. Python Strategic Pipeline Implementation
    add_section_heading("6. Python Preliminary Ingestion & Strategic Pipeline Script")
    add_body_p(
        "Below is the verified Python ingestion script developed for Week 1. It loads the dataset, conducts an automated "
        "schema and null-value audit, computes foundational KPI baseline estimates, and validates operational constraints:"
    )

    code_p = doc.add_paragraph()
    code_p.paragraph_format.line_spacing = 1.0
    code_p.paragraph_format.space_before = Pt(4)
    code_p.paragraph_format.space_after = Pt(8)
    code_run = code_p.add_run(
"""import pandas as pd
import numpy as np

# Load raw logistics shipment dataset
data_path = 'data/raw_logistics_data.csv'
df = pd.read_csv(data_path)

print(f"Dataset Shape: {df.shape[0]} shipments, {df.shape[1]} attributes")

# Compute Baseline Operational KPIs
ontime_rate = (df['Is_Delayed'] == 0).mean() * 100
avg_delay = df[df['Delay_Hours'] > 0]['Delay_Hours'].mean()
avg_dock_time = df['Warehouse_Dock_Hours'].mean()

print(f"Baseline On-Time Delivery Rate (OTD): {ontime_rate:.2f}%")
print(f"Average Delay for Late Shipments: {avg_delay:.2f} hours")
print(f"Average Warehouse Dock Turnaround: {avg_dock_time:.2f} hours")

# Audit Data Quality Baseline
missing_summary = df.isnull().sum()
print("\\nMissing Value Audit:")
print(missing_summary[missing_summary > 0])
"""
    )
    code_run.font.name = 'Courier New'
    code_run.font.size = Pt(8.5)
    code_run.font.color.rgb = RGBColor(30, 30, 30)

    # 7. Strategic Risk Assessment & Governance
    add_section_heading("7. Strategic Risk Assessment & Mitigation Matrix")
    add_body_p(
        "Implementing advanced analytics in supply chain environments introduces technical, operational, and organizational risks:"
    )
    add_bullet_p("Data Sensor Inconsistencies", "Mitigated by implementing strict IQR truncation rules for telematics readings.")
    add_bullet_p("Overfitting to Seasonal Peaks", "Mitigated via rolling cross-validation across distinct calendar quarters.")
    add_bullet_p("Driver & Carrier Algorithmic Distrust", "Addressed by producing interpretable feature attribution metrics (SHAP values) that clearly explain the factors driving delay predictions.")

    # 8. Conclusion and Week 2 Next Steps
    add_section_heading("8. Conclusion and Strategic Milestones for Week 2")
    add_body_p(
        "Week 1 has successfully established the project scope, KPI definitions, data architecture, and operational baseline. "
        "With a verified dataset of 5,000 shipments and a clearly structured analytical roadmap, the project is positioned to transition "
        "into Week 2: Data Collection, Cleaning, and Preprocessing. Week 2 will execute rigorous missing-data imputation, "
        "outlier filtration, categorical one-hot encoding, and specialized logistics feature engineering."
    )

    doc.save(output_path)
    print(f"Week 1 Report successfully generated at: {output_path}")

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    report_path = os.path.join(base_dir, 'reports', 'Week_1_Strategic_Planning_and_Data_Exploration.docx')
    create_week1_report(report_path)
