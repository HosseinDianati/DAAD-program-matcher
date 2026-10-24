# DAAD Program Matcher & Admissions Advisor

Autonomous German university study program matcher and admissions advisor for international students and educational consultants.

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
5. Generates a production-ready `.xlsx` spreadsheet matching the client template.

---

## Installation

### In Gemini / Claude Code (One Command)

Run this directly in chat:
```bash
/plugin marketplace add HosseinDianati/daad-program-matcher
/plugin install daad-program-matcher@daad-program-matcher
```

Or simply paste this repo link into the chat:
```text
https://github.com/HosseinDianati/daad-program-matcher
install skill
```

---

## Python Dependencies
```bash
pip install -r requirements.txt
```
