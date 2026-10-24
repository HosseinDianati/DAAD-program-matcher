---
name: daad-program-matcher
description: Autonomous German university study program matcher and admissions advisor. Parses candidate CVs (PDF/text/voice), computes German GPA equivalencies (Modified Bavarian Formula), deeply explores the DAAD International Programmes database, verifies university examination regulations, and outputs standardized 11-column Excel workbooks and Google Sheets. Triggers on DAAD, German university search, study in Germany, match CV to German programs, Uni-Assist programs, Winter/Summer intake search.
---

# DAAD Program Matcher & Admissions Advisor Skill

## Overview
This skill guides the agent to match international students to German higher education study programs using the official **DAAD International Programmes** database and official university examination regulations (*Zugangssatzung / Prüfungsordnung*).

The output strictly follows the 11-column standard established in `Winter search list- Javad Mahmoodi.xlsx` and `Winter search Template.xlsx`:
1. `University`
2. `Program`
3. `Tuition fee`
4. `Semester fee`
5. `Intake`
6. `Deadline`
7. `Requirements`
8. `Application Method`
9. `Program link`
10. `Note`
11. `Apply Link`

---

## Stage 1: Candidate Profile Ingestion & German Academic Translation

### 1. Grade Conversion (Modified Bavarian Formula)
German grades range from **1.0 (Best)** to **4.0 (Minimum Passing)**. 5.0 is Fail.
Formula:
$$N = 1 + 3 \times \frac{N_{\max} - N_d}{N_{\max} - N_{\min}}$$
Where:
- $N$: German GPA equivalent
- $N_{\max}$: Maximum achievable grade in candidate's home country
- $N_{\min}$: Minimum passing grade in candidate's home country
- $N_d$: Candidate's obtained grade

**Common Grading Scales:**
- **Iran (Bachelor)**: $N_{\max} = 20.0$, $N_{\min} = 10.0$ (e.g., $18.05 \rightarrow 1.58$ German GPA).
- **Iran (Master)**: $N_{\max} = 20.0$, $N_{\min} = 12.0$.
- **India (10-point CGPA)**: $N_{\max} = 10.0$, $N_{\min} = 4.0$.
- **US / 4.0 scale**: $N_{\max} = 4.0$, $N_{\min} = 2.0$.

### 2. Country-Specific Mandatory Credentials
- **APS Certificate (Akademische Prüfstelle)**:
  - **MANDATORY** for applicants with university degrees from **Iran, India, China, and Vietnam**.
  - Always flag in `Note` or `Requirements` if the candidate is from one of these countries.
- **Language Requirements**:
  - **English-taught**: Typically requires IELTS Academic $\ge 6.5$ (or TOEFL iBT $\ge 90$). If a candidate has IELTS 7.5 (C1 level), they meet virtually all English requirements.
  - **German-taught / Bilingual**: Check whether German A1 is strictly required upfront or can be completed during the first academic year (e.g., University of Passau, University of Bayreuth).

---

## Stage 2: DAAD Deep Discovery & Scraping

### 1. DAAD Solr API Endpoint
Query the DAAD International Programmes database:
```
GET https://www2.daad.de/deutschland/studienangebote/international-programmes/api/solr/en/search.json
Parameters:
  q: search term (e.g. "architecture", "energy", "computer science")
  degree[]: 2 (Master's degree)
  lang[]: 2 (English only) or 1 (German) or 3 (Both)
  fee[]: 0 (No tuition fees) if student requested tuition-free
  limit: 20
  offset: 0
  sort: 4
```

### 2. Detail Page Inspection
For every shortlisted program, visit its DAAD detail page (`/detail/{id}/#tab_registration` and `#tab_costs`):
- **Tuition fees**: Scrape exact amount per semester (or `none`).
- **Semester contribution**: Scrape administrative/ticket fee (e.g., `300 Euro per semester`).
- **Beginning / Intake**: `Winter semester` or `Summer and Winter`.
- **Application periods / Deadlines**: Non-EU international deadlines (e.g., `15 December - 15 March`, `1 April - 31 May`).
- **Academic admission requirements**: Extract degree prerequisite, minimum ECTS distribution, and portfolio/work experience requirements.
- **Language requirements**: IELTS/TOEFL minimum scores, German CEFR levels and waivers.
- **Submit application via**: Identify `uni-assist` vs university direct portal.

### 3. University Regulation Verification
- Follow the **"Course website"** link to verify:
  - Exact ECTS breakdown by subject (e.g., *at least 35 CP in fundamentals of mathematics, at least 20 CP in programming*).
  - Preliminary Review Documentation (**VPD**) requirement via Uni-Assist.
  - Direct application portal link (`Apply Link`).

---

## Stage 3: The 11-Column Output Specification

| Column | Header | Exact Format Guidelines |
|---|---|---|
| A | `University` | Full official name in English (e.g., `HafenCity University Hamburg`, `TU Dresden`) |
| B | `Program` | Full official degree name (e.g., `Resource Efficiency in Architecture and Planning (REAP) (M.Sc.)`) |
| C | `Tuition fee` | `none` or exact amount (e.g., `1500 Euro per Semester`, `850 EUR`) |
| D | `Semester fee` | e.g., `340 Euro per semester`, `120 Euro per semester` |
| E | `Intake` | `Winter` or `Summer and Winter` |
| F | `Deadline` | Exact window for Non-EU/international: e.g., `1 June to 15 July`, `15 March`, `2026-05-31` |
| G | `Requirements` | Numbered list covering: 1. Bachelor's discipline & ECTS subject breakdown; 2. Minimum GPA / German grade cutoff; 3. English language score (IELTS/TOEFL); 4. German requirement or 1st-year waiver; 5. GRE/Portfolio/VPD details |
| H | `Application Method` | `Uni-Assist`, `Uni Portal`, or `Uni Portal (via VPD)` |
| I | `Program link` | Direct link to the university course curriculum or DAAD registration tab |
| J | `Note` | Critical caveats (e.g., "APS Certificate required for Iranian/Indian degrees", "Portfolio required", "Baden-Württemberg €1500/sem tuition applies") |
| K | `Apply Link` | Direct link to the application login portal (`https://my.uni-assist.de/`, `CampusOnline`, `HISinOne`, `VIBS`, `AlmaWeb`) |

---

## Stage 4: Execution Scripts

- **Grade calculation**: `python C:\Users\hosse\.gemini\config\skills\daad-program-matcher\scripts\bavarian_calculator.py`
- **Excel generation**: `python C:\Users\hosse\.gemini\config\skills\daad-program-matcher\scripts\excel_generator.py`
