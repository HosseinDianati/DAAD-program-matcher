---
name: daad-program-matcher
description: Autonomous German university study program matcher and admissions advisor. Parses candidate CVs (PDF/text/voice), computes German GPA equivalencies (Modified Bavarian Formula), deeply explores the DAAD International Programmes database, performs two-phase Google search verification for department-specific deadlines and URLs, extracts complete DAAD document checklists, objectively determines admission chance tiers (Safe/Target/Reach), and outputs standardized 13-column Excel workbooks, Google Sheets matrices, and interactive client web dashboards. Triggers on DAAD, German university search, study in Germany, match CV to German programs, Uni-Assist programs, Winter/Summer intake search.
---

# DAAD Program Matcher & Admissions Advisor Skill

## Overview
This skill guides the agent to match international students to German higher education study programs using the official **DAAD International Programmes** database and official university examination regulations (*Zugangssatzung / Prüfungsordnung*).

The output strictly follows the production **13-column standard**:
1. `University`
2. `Program`
3. `Admission Chance` (`🟢 High (Safe)`, `🟡 Medium (Target)`, `🔴 Low (Reach)`)
4. `Tuition fee`
5. `Semester fee`
6. `Intake`
7. `Deadline` (verified department-specific non-EU application cutoff)
8. `Requirements` (ECTS subject distributions, German/English CEFR levels, GRE/portfolio)
9. `Required Documents` (complete checklist of application paperwork)
10. `Application Method` (`Uni-Assist`, `Uni Portal`, or `Uni Portal (via VPD)`)
11. `Program link` (verified canonical departmental program page)
12. `Note` (APS certificate flags, state tuition caveats, aptitude tests)
13. `Apply Link` (direct link to the application submission/login portal)

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
  - Always flag in `Note` and include in `Required Documents` if candidate is from these countries.

---

## Stage 2: Two-Phase Search & Verification Pipeline

### Phase A: DAAD Discovery & Initial Candidate Compilation
1. Query the DAAD Solr API:
   `https://www2.daad.de/deutschland/studienangebote/international-programmes/api/solr/en/search.json`
2. Fetch program details (`#tab_overview`, `#tab_costs`, `#tab_registration`).
3. Compile the preliminary candidate list into memory.

### Phase B: Department-Level Google/Web Verification
DAAD links often redirect to generic university homepages or cite general nationwide deadlines (e.g. July 15) rather than the strict departmental cutoff. For each shortlisted program:
1. Search via Google / Search tool:
   `"<University Name>" "<Program Name>" admission requirements deadline international Non-EU`
2. **Correct the canonical program link**: Verify that `Program link` leads to the specific faculty/department degree overview (not the generic university homepage).
3. **Verify exact faculty non-EU deadline**: Identify whether the department closes earlier for non-EU students (e.g. March 15, April 30, May 15, May 31).
4. **Verify direct application portal**: Distinguish between `my.uni-assist.de`, institutional HISinOne/QIS, CampusOnline, Primuss, or MoveOn.

---

## Stage 3: Admission Chance Categorization Engine

Every program must be consistently categorized into one of three tiers based on objective academic matching:

1. **`High (Safe)`**:
   - Candidate's German GPA is significantly stronger than the minimum cutoff (e.g., candidate has $\le 1.8$, requirement is $2.5$ or open admission / *zulassungsfrei*).
   - Candidate satisfies 100% of prerequisite ECTS credit distributions in core subjects.
   - English/German language scores exceed minimum thresholds.
2. **`Medium (Target)`**:
   - Candidate fulfills all formal academic criteria, but admission is restricted (*Numerus Clausus* / NC) or subject to a competitive aptitude review (*Eignungsverfahren*), portfolio assessment, or interview.
   - Candidate GPA is competitive but within typical NC cutoffs.
3. **`Low (Reach)`**:
   - Highly selective quotas for non-EU applicants (e.g. TU Munich, RWTH Aachen).
   - Prerequisite credit gaps in specific advanced modules (e.g., missing 10 ECTS in higher mathematics or theoretical computer science).
   - Rigid portfolio or interview cutoffs with low international acceptance rates.

---

## Stage 4: Required Documents Extraction Protocol

For each program, populate Column 9 (`Required Documents`) with an itemized checklist derived from DAAD and university regulations:
1. Official Bachelor's Degree Certificate (original + certified German/English translation).
2. Complete Academic Transcript of Records with official grading scale explanation.
3. **APS Certificate** (mandatory for Iran, India, China, Vietnam).
4. **Module Handbook / Course Syllabi** (crucial for German faculty credit equivalency checks).
5. Official Language Certificate (IELTS Academic, TOEFL iBT, or TestDaF/Goethe).
6. Tabular Curriculum Vitae (CV) in European/Europass format.
7. Letter of Motivation / Statement of Purpose (incorporating any department-specific prompts or word limits).
8. Work Portfolio (specifying format, max pages, dimensions DIN A3/A4, e.g. for Architecture/Design).
9. Proof of Professional Work Experience / Internship (specifying duration in months).
10. **Uni-Assist VPD (Vorprüfungsdokumentation)** if applying through university portal with pre-evaluation.

---

## Stage 5: Output Generation (Excel, Google Sheet, and Interactive Web Portal)

1. **Excel Generation**:
   Execute `scripts/excel_generator.py` to create the formatted 13-column `.xlsx` workbook with auto-wrapping, bold headers, and color-coded chance cells.
2. **Google Sheets Export**:
   Generate a TSV/matrix dataset ready for direct pasting or syncing via the Google Sheets integration.
3. **Interactive Client Dashboard**:
   Execute `scripts/dashboard_generator.py` to produce `client_portal.html` containing:
   - Dynamic Add Program & Remove Program buttons.
   - Filter by Admission Chance (High/Medium/Low).
   - In-browser 1-click download of the updated Excel file.
