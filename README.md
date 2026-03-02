# 🚌 Multi-Criteria Decision Making Tool for Mass Transportation Using AHP

> **A Hybrid AHP–WSM (Analytic Hierarchy Process – Weighted Sum Model) decision support system to identify the best bus allocation for public transport routes in the Tiruppur region, Tamil Nadu.**

---

## 👥 Team

| Name | Role |
|---|---|
| Raja Gopal R | Lead Developer & AHP Modelling |
| Sasvat Baalan A | Data Collection & Validation |
| Rahman Sherief A | Analysis & Presentation |

---

## 📌 Project Overview

Selecting the best bus for a transport route involves multiple competing criteria — travel time, passenger comfort, fare cost, and service frequency. This tool uses the **Hybrid AHP–WSM methodology** to systematically rank **100 real TNSTC/SETC buses** operating from Tiruppur, Tamil Nadu, based on expert judgment and real route data.

The system is designed for transport planners, researchers, and policymakers who need a **data-driven, consistent, and transparent** method to prioritize bus allocations.

---

## ✨ Features

- 🧠 **AHP (Analytic Hierarchy Process)** — calculates criteria weights from expert pairwise comparisons with consistency checking (CI/CR)
- 📊 **WSM (Weighted Sum Model)** — scores and ranks all 100 bus alternatives
- 🌐 **Interactive Browser UI** — expert fills in Saaty scale comparisons visually; no manual Excel editing needed
- 🔗 **Fully Connected Workflow** — one command opens the browser, captures input, runs analysis, and saves results
- 📋 **Professional Excel Output** — formatted Results sheet with sections for weights, consistency check, normalized data, and final ranking
- 🚌 **Real TNSTC/SETC Data** — 100 buses across 7 route categories sourced from actual Tamil Nadu bus services

---

## 🗂️ Project Structure

```
📁 ahp-bus-allocation/
│
├── newcr.py           # Main script — runs full AHP-WSM pipeline
├── expert_tool.py     # Flask web app — expert pairwise input UI
├── ahp_input.xlsx     # Excel workbook with all input data & results
└── README.md          # This file
```

### Excel Sheets

| Sheet | Description |
|---|---|
| `Criteria` | Pairwise comparison matrix (written by expert tool) |
| `Alternatives_Data` | 100 real TNSTC/SETC buses with TravelTime, Comfort, Cost, Frequency |
| `Criteria_Config` | Benefit/Cost classification for each criterion |
| `Expert_Input` | Copy of the expert comparison matrix for reference |
| `Results` | Final ranked output with all 4 sections |

---

## 🔢 Methodology

### Phase 1 — AHP (Criteria Weighting)

1. Expert fills pairwise comparisons using the **Saaty 1–9 scale**
2. Comparison matrix is normalized column-wise
3. Priority vector (weights) = row means of normalized matrix
4. Consistency is verified: **CR < 0.10** required

$$CR = \frac{CI}{RI}, \quad CI = \frac{\lambda_{max} - n}{n - 1}$$

### Phase 2 — WSM (Alternative Scoring)

1. Alternatives normalized by criterion type:
   - **Benefit** (higher = better): $r_{ij} = \frac{x_{ij}}{x_j^{max}}$
   - **Cost** (lower = better): $r_{ij} = \frac{x_j^{min}}{x_{ij}}$
2. Final score = weighted sum: $S_i = \sum_{j=1}^{n} w_j \cdot r_{ij}$
3. Buses ranked by descending final score

### Criteria

| Criterion | Type | Description |
|---|---|---|
| TravelTime | Cost | Journey duration in minutes (lower = better) |
| Comfort | Benefit | Passenger comfort rating 1–10 (higher = better) |
| Cost | Cost | Fare in ₹ (lower = better) |
| Frequency | Benefit | Buses per hour (higher = better) |

---

## 🚀 How to Run

### Prerequisites

```bash
pip install numpy pandas openpyxl flask
```

### Run (One Command Does Everything)

```bash
python3 newcr.py
```

**What happens automatically:**

```
Step 1 → Browser opens at http://127.0.0.1:5050
         Expert fills pairwise comparisons
         Expert clicks "Save to Excel"

Step 2 → Terminal detects save
         Reads updated criteria matrix
         Normalizes all 100 bus alternatives
         Calculates AHP-WSM scores
         Saves ranked results to Excel → Results sheet
```

### Run Expert Tool Standalone (optional)

```bash
python3 expert_tool.py
```

---

## 📊 Sample Output

```
=== CRITERIA WEIGHTS (AHP) ===
TravelTime    0.5579   (55.8%)
Comfort       0.2633   (26.3%)
Cost          0.1219   (12.2%)
Frequency     0.0569    (5.7%)

CI=0.0395,  CR=0.0439  ✔ CONSISTENT

=== TOP 5 BEST BUS ALLOCATIONS ===
 1. Bus_084  Tiruppur–Coimbatore (Volvo AC Seater)     Score=0.7952
 2. Bus_056  Tiruppur–Coimbatore (Point-to-Point AC)   Score=0.7592
 3. Bus_098  Tiruppur–Avinashi (Superfast AC)          Score=0.6778
 4. Bus_092  Tiruppur City (Economy)                   Score=0.6734
 5. Bus_091  Tiruppur City MTC-type (Economy)          Score=0.6625
```

---

## 🗃️ Dataset — 100 Real TNSTC/SETC Buses

Real bus data from **Tiruppur, Tamil Nadu** covering:

| Category | Count | Routes |
|---|---|---|
| Short routes (30–90 min) | 20 | Coimbatore, Erode, Avinashi, Palladam, Kangeyam |
| Medium routes (120–240 min) | 20 | Salem, Ooty, Palani, Karur, Trichy, Dindigul |
| Long routes (240–480 min) | 20 | Chennai, Bangalore, Madurai, Guruvayur, Thrissur |
| SETC Interstate | 10 | Tiruchendur, Kanyakumari, Velankanni, Rameswaram |
| Semi-urban | 10 | Gobichettipalayam, Bhavani, Sathyamangalam |
| Premium (Volvo / Scania) | 10 | AC Sleeper, Volvo AC Seater |
| Intra-city / Feeder | 10 | City buses, Industrial area shuttles |

**Fare range:** ₹8 (city bus) to ₹850 (Scania AC Sleeper to Chennai)
**Service types:** Ordinary Express, Deluxe, Ultra Deluxe, SFS AC, Economy AC (3×2), Volvo AC, Scania AC Sleeper

---

## 🛠️ Tech Stack

| Component | Technology |
|---|---|
| Core Algorithm | Python (NumPy, Pandas) |
| Excel I/O | openpyxl |
| Web UI | Flask + Vanilla JavaScript |
| Data Input | HTML with Saaty Scale UI |
| Output | Formatted Excel (.xlsx) |

---

## 📖 References

- Saaty, T.L. (1980). *The Analytic Hierarchy Process*. McGraw-Hill.
- TNSTC Official Bus Schedules — tnstc.in
- SETC Tamil Nadu — setc.tn.gov.in
- Hwang, C.L. & Yoon, K. (1981). *Multiple Attribute Decision Making*. Springer.

---

## 📄 License

This project is developed for academic purposes as part of a Multi-Criteria Decision Making study on mass transportation systems in Tamil Nadu.

---

*Made with ❤️ in Tiruppur, Tamil Nadu 🇮🇳*
