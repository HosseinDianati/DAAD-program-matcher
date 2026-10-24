# DAAD Program Matcher & Admissions Advisor

Autonomous German university admissions matcher for educational consultants and prospective students.

## What This Skill Does
1. Parses student CVs (PDF, images, text, or voice transcripts).
2. Calculates German GPA equivalence using the Modified Bavarian Formula:
   $$N = 1 + 3 \times \frac{N_{\max} - N_d}{N_{\max} - N_{\min}}$$
3. Checks mandatory country credentials (e.g. APS Certificate for Iran, India, China, Vietnam).
4. Deeply searches the official DAAD database and scrapes official university examination regulations for:
   - Credit point / ECTS subject distributions.
   - Non-EU application deadlines.
   - Tuition fees and semester ticket fees.
   - Exact application method (Uni-Assist vs Uni Portal vs VPD).
   - Direct links to both the course curriculum and the application registration portal.
5. Generates a production-ready `.xlsx` spreadsheet matching the standard 11-column template.

---

## Team Setup Instructions (1 Minute)

### Step 1: Install Python Dependencies
Open Terminal / PowerShell and run:
```bash
pip install openpyxl beautifulsoup4 requests
```

### Step 2: Copy the Skill Folder
Place the `daad-program-matcher` folder into your global Antigravity/Gemini skills directory:

* **Windows:**
  `C:\Users\<YourUsername>\.gemini\config\skills\daad-program-matcher`
* **macOS / Linux:**
  `~/.gemini/config/skills/daad-program-matcher`

*(Alternatively, if your team shares a project folder or Git repo, simply place this folder inside your repo at `.gemini/skills/daad-program-matcher` — anyone opening that folder will have the skill active automatically!)*

---

## How to Use

Simply start a conversation with the assistant and attach a CV or give your instructions:
- *"Match this attached PDF CV to accredited English-taught Master's programs in Germany for Winter semester."*
- *"Find tuition-free German universities in Computer Science for a student with a 3.2 US GPA and IELTS 7.0."*
- *"Here is a voice note of student preferences. Output the results using our Winter search template."*
