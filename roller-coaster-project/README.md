# Roller Coaster Polynomial Project (Algebra 2)

Students design a roller coaster and model it with a polynomial (time in seconds vs. height in feet).
A Google Form + Google Sheet checker automatically verifies each student's work.

| File | What it is |
|---|---|
| `Roller_Coaster_Project_Revised.docx` | Student handout: requirements, 7 deliverables, typing guide for the form, rubric (100 pts) |
| `Coaster_Checker_Teacher_Guide.docx` | Teacher setup guide, sample test submissions, how to read the results |
| `CoasterChecker.gs` | Google Apps Script that builds the student Form and checks every submission |

## What the checker does

- Expands each student's **factored (intercept) form** into **standard form** and compares it term by term with the standard form they typed.
- Checks the **a-value** against the point the student used (usually the platform, `(0, height)`).
- Checks the requirements: ride ≥ 120 s, ≥ 3 turning points during the ride, a tunnel (y < 0), a platform above ground, realistic heights.
- Checks the student's answers for end behavior, turning point times, and range.
- Builds an answer key for each student's coaster: turning points, intervals of increase/decrease, domain, range.

## Quick setup

1. Create a blank Google Sheet.
2. **Extensions → Apps Script**, paste in `CoasterChecker.gs`, then save.
3. Refresh the Sheet, then choose **Coaster Checker → 1. Create student form**, and authorize the script.
4. Post the student link it shows in Google Classroom.

Requirements and tolerances can be changed on the **Settings** tab, then **Coaster Checker → Re-check all submissions**.
See the teacher guide for details.
