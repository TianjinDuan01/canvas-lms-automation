# Canvas TA Automation Toolkit

A collection of Python scripts to automate common TA workflows for courses hosted on Canvas LMS. Built as part of TA automation infrastructure at Johns Hopkins University, but designed to be reusable for any course.

---

## Overview

| Script | Purpose |
|--------|---------|
| `canvas_discussion_grader.py` | Grade discussion board participation (posts + replies) across multiple rounds |
| `a2_anonymize_submissions.py` | Validate and anonymize student HTML game submissions for peer evaluation |
| `a3_process_submissions.py` | Download, validate, and organize job application submissions across multiple attempts |
| `a3_grading.py` | Scrape leaderboard rankings and calculate final scores using a grading formula |

---

## Prerequisites

### Python
Python 3.8 or higher.

### Dependencies
```bash
pip install requests openpyxl beautifulsoup4
```

### Canvas API Token
All scripts authenticate via the Canvas REST API using a personal access token.

**To generate a token:**
1. Log in to Canvas → click your profile icon (top right) → **Account** → **Settings**
2. Scroll to **Approved Integrations** → click **New Access Token**
3. Enter a purpose (e.g. `TA Grading Scripts`) and an expiration date
4. Copy the token — **it is only shown once**

> ⚠️ Keep your token private. Never commit it to version control. The scripts will prompt you to enter it at runtime, or you can set it as an environment variable.

**Finding your Course ID and other IDs:**
All IDs are visible in Canvas URLs:
```
https://your-institution.instructure.com/courses/COURSE_ID/discussion_topics/DISCUSSION_ID
                                                        ↑ Course ID         ↑ Discussion ID
```

---

## How the Canvas API Works

These scripts use the [Canvas LMS REST API](https://canvas.instructure.com/doc/api/). All requests are authenticated with a Bearer token in the HTTP header:

```python
headers = {"Authorization": f"Bearer {YOUR_TOKEN}"}
response = requests.get(url, headers=headers)
```

**Key endpoints used:**

| Endpoint | What it returns |
|----------|----------------|
| `/api/v1/courses/:id/enrollments` | All students, TAs, and instructors with their roles |
| `/api/v1/courses/:id/discussion_topics/:id/entries` | Top-level posts in a discussion |
| `/api/v1/courses/:id/discussion_topics/:id/entries/:id/replies` | Replies to a specific post |
| `/api/v1/courses/:id/assignments/:id/submissions` | All student submissions for an assignment |
| `/api/v1/courses/:id/quizzes/:id/submissions` | All student quiz submissions |

**Pagination:** Canvas returns results in pages of up to 100 items. All scripts handle pagination automatically by following the `Link: rel="next"` header.

**Role filtering:** The API returns everyone enrolled in the course. Scripts automatically skip instructors, TAs, and designers using the `type` field in enrollment responses.

---

## Scripts

### 1. `canvas_discussion_grader.py`

Grades student participation across multiple Canvas discussion rounds. For each student, it checks whether they made an original post and a reply in each round, then calculates a score.

**Grading logic:**
- Each post = 1 point
- Each reply = 1 point
- Late submissions (after deadline) are flagged but still counted — manual review recommended

**Configuration (top of file):**
```python
CANVAS_BASE_URL = "https://your-institution.instructure.com"
COURSE_ID       = "YOUR_COURSE_ID"
DISCUSSION_IDS  = ["ID1", "ID2", "ID3"]  # one per round
DEADLINE_UTC    = datetime(2026, 6, 1, 3, 59, 0, tzinfo=timezone.utc)  # adjust as needed
```

**Run:**
```bash
python3 canvas_discussion_grader.py
```

**Output:** `discussion_grades_TIMESTAMP.csv`

| Column | Description |
|--------|-------------|
| Student Name | Canvas display name |
| D1 Post / D1 Reply | Yes / No |
| D1 Post Late / D1 Reply Late | LATE if submitted after deadline |
| Total Posts / Total Replies | Count across all rounds |
| Score (out of N) | Total points |
| Any Late | LATE if any submission was after deadline |

---

### 2. `a2_anonymize_submissions.py`

Two-step pipeline for anonymizing student HTML game submissions before peer evaluation.

**Requires two Excel files:**
- `Public_Assigned_Names.xlsx` — maps each student to their assigned SubmissionID (e.g. `xaji0y`)
- `PRIVATE_Mapping.xlsx` — maps each SubmissionID to an anonymous GameID (e.g. `022`) — **keep private**

**Step 1 — Validate filenames (run before deadline):**
Checks that each student submitted:
- `{SubmissionID}.html` (required)
- `{SubmissionID}_prompt.txt` (required)
- `{SubmissionID}_title.txt` (optional)

Outputs a report of students who need to fix their filenames.

**Step 2 — Download and anonymize (run after Step 1 is clean):**
Downloads the latest HTML submission from each student and renames it from `xaji0y.html` → `022.html` per the private mapping. Output folder is ready to upload to the course website.

**Run:**
```bash
python3 a2_anonymize_submissions.py
# Then choose: 1 (validate) or 2 (download + anonymize)
```

---

### 3. `a3_process_submissions.py`

Downloads and organizes job application submissions (`.docx` files) for a multi-attempt assignment. Handles Canvas auto-suffixes (e.g. `resume-2.docx`) and validates filenames against a required format.

**Expected filename format:**
```
JHED_JobID_mask.docx
e.g. jsmith99_001_batman123.docx
```

**Run:**
```bash
python3 a3_process_submissions.py
# Then choose which attempt to process (1, 2, or 3)
```

**Output per run:**
```
a3_submissions/
  Attempt1/
    001/
      mask-Attempt1.docx
    002/
      ...
a3_mapping_Attempt1_TIMESTAMP.xlsx   ← PRIVATE — one sheet per Job ID
a3_warnings_Attempt1_TIMESTAMP.txt   ← students with invalid filenames
```

---

### 4. `a3_grading.py`

Scrapes a leaderboard website to retrieve AI-evaluated Mean Ranks for each submission, then calculates final scores.

**Grading formula:**
```
Score = 17 - (r - 1) × (9 / (n - 1))
```
where `r` = best Mean Rank across all attempts, `n` = total applicants for that job.
- Ranked 1st → 17 points (full marks)
- No submission or format violation → 0 points

**Run:**
```bash
python3 a3_grading.py
# Enter leaderboard URL and paths to a3_mapping Excel files
```

**Output:** `a3_grades_TIMESTAMP.xlsx`
- **Grades sheet** — one row per student with all attempts, best rank, and final score
- **Leaderboard Data sheet** — raw scraped rankings for reference

---

## Gradebook Import

All grading scripts produce output that can be imported directly into the Canvas Gradebook.

**Format required by Canvas:**
```csv
Student,ID,SIS Login ID,Section,Assignment Name
"Smith, John",12345,jsmith1@jh.edu,Section 1,5
```

**To import:**
1. Canvas Gradebook → **Import** (top right)
2. Upload the CSV
3. Preview and confirm

> The assignment column name must exactly match the assignment name in Canvas (including capitalization and spaces).

---

## Privacy and Security Notes

- API tokens have the same access level as your Canvas account. **Never share or commit them.**
- Student data (names, grades, submissions) is protected under FERPA. Keep all output files within your institution's approved systems.
- The `PRIVATE_Mapping.xlsx` file in the A2 pipeline maps anonymous game IDs to real student identities — treat it as confidential and do not distribute to students.
- Tokens can be revoked at any time: Canvas → Account → Settings → delete the token.

---

## Adapting for Your Course

1. Update `CANVAS_BASE_URL` to your institution's Canvas URL
2. Update `COURSE_ID` and relevant assignment/discussion IDs
3. Adjust deadlines and grading formulas as needed
4. The scripts are modular — each one is independent and can be used without the others

---

## Author

Built by Elina (Tianjin Duan) as part of TA automation infrastructure
