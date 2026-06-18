# ☕ Automated Cafe Chain Operations & Sales Analytics Pipeline (iPOS ETL)

An end-to-end automated **ETL** (Extract, Transform, Load) pipeline designed for cafe chain owners and store managers to optimize daily operations, control inventory, track staff performance, and drive data-backed business decisions.

This project automates the extraction of daily sales logs, transaction invoices, customer data, and item preparation times from the **iPOS (Fabi)** system. Raw data is cleaned, structured via Python, and automatically pushed to **Google Sheets** to feed real-time **Looker Studio (Data Studio)** business intelligence dashboards.

---

## 🎯 Key Business & Operational Impacts

This pipeline translates raw transactional logs into two specialized dashboards to assist management in three core pillars:

1. **Operational Oversight & Decisions:** Monitoring peak hours, product velocity, and daily revenue trends to balance store workflows and manage supply chains effectively.
2. **Inventory & Cost Control:** Merging sales volumes with inventory logs to flag stock discrepancies, minimize wastage, and optimize reorder points.
3. **Staff Auditing & Performance:** Tracking individual server metrics (sales volume, ticket sizes, combo upsells) and monitoring order anomalies (voids, modifications) to minimize internal fraud.

---

## 🏗️ System Architecture

```text
 [ iPOS (Fabi) System ]
          │
          ▼ (Extract)
 ┌─────────────────────────────────────────────────────────────┐
 │ Python Crawl Scripts (Triggered daily via taskschd.msc)     │
 └─────────────────────────────────────────────────────────────┘
          │
          ▼ (Raw File Storage)
 ┌─────────────────────────────────────────────────────────────┐
 │ Data Transformation Engine (Employee_Report.py)             │
 │  — Cleaned & joined sales, invoice, and preparation logs     │
 │  — Aggregated KPI matrix (Sales, Inventory, Staff Metrics)  │
 └─────────────────────────────────────────────────────────────┘
          │
          ▼ (Load via Google Sheets API)
 [ Google Sheets Database ]
          │
          ▼ (Live Connection)
 [ Looker Studio BI Dashboard ] ──► End-user Report (Managers/Owners)
```
---

## 📂 Repository Structure

* `data/` - Raw data snapshots crawled from iPOS and Cleaned, structured data ready for Google Sheets
* `reports/` - Previews and static exports of the BI dashboards
* `scripts/` - Core ETL and Scraping engine
* `.gitignore` - Safeguards API keys and local Excel snapshots
* `README.md` - Project documentation
* `requirements.txt` - Python dependencies

---

## 🛠️ Data Pipeline & ETL Technical Details

1. **Extract**
* Production-grade custom Python scrapers (`scripts/crawl_*.py`) securely request daily transactional files from the iPOS endpoint.
* Scheduled autonomously every midnight using the **Windows Task Scheduler** (`taskschd.msc`) to capture the previous day's data without any manual oversight.
2. **Transform**
* Data Cleansing: Standardizes conflicting datetime fields across multi-tables into `%Y-%m-%d %H:%M:%S`, drops broken records, and unifies product naming conventions.
* Feature Engineering: Maps invoice timestamps against product categories to compute preparation velocities and links modification logs to specific cashier profiles.
3. **Load**
* Connects seamlessly via OAuth2 (`google-api-python-client`) to overwrite or append the latest structured matrices into individual worksheets on Google Sheets.

---

##  📊 Business Intelligence Dashboards (Looker Studio)

The system updates two tailored interactive reporting dashboards:

**Dashboard 1: Sales, Customer Loyalty (CRM) & Store Inventory**
* Sales Analytics: Track net revenue, Average Order Value (AOV), and peak traffic windows to optimize store operating hours.
* Loyalty & CRM: Measure new customer acquisition vs. returning customer retention rates ($%$) along with their preferred product categories.
* Logistics & Inventory: Connect sales velocity with stock consumption to monitor raw material usage, highlight variances, and flag stockouts.
  [See static preview reports/sales & inventory dashboard.pdf](https://github.com/thuhang24ld/data-analysis-projects/blob/main/reports/sales%20%26%20inventory%20dashboard.pdf)
  
**Dashboard 2: Employee Performance & Operational Audit**
* Staff Performance: Breakdown revenue, successful checkouts, and promotional/combo upsell success rates by individual server.
* Risk & Fraud Mitigation: Deep-dive into order change logs to flag employee accounts with unusually high item cancellations, voided bills, or manual price overrides.
  [See static preview: reports/employee dashboard.pdf](https://github.com/thuhang24ld/data-analysis-projects/blob/main/reports/employee%20dashboard.pdf)

--- 
## ⚙️ Getting Started & Setup

**Prerequisites**
* Python 3.9+
* Google Cloud Platform (GCP) account with Google Sheets API enabled.
* Service Account Credentials saved as a JSON key.
  
**Installation**
1. Clone the repository:
```Bash
git clone [https://github.com/your-username/automated-cafe-analytics.git](https://github.com/your-username/automated-cafe-analytics.git)
cd automated-cafe-analytics
```
2. Install required dependencies:
```Bash
pip install -r requirements.txt
```
3. Save your GCP service account JSON key to the root directory as credentials.json (ensure it is listed in your .gitignore).

**Automation Scheduling (Windows)**
1. Open Task Scheduler (`taskschd.msc`).
2. Create a Basic Task -> Set Trigger to Daily (e.g., 12:30 AM).
3. Action: Start a Program.
4. Program/Script: Path to your local `python.exe`.
5. Add arguments: `scripts/crawl_sale_by_date.py` (Configure separate tasks or a single runner script to execute all scrapers sequentially).
