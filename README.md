# Food-Safety-Practice-Questions
Mock test strictly based on our study guideline in NUTR248, Simmons University

Practice for the **ServSafe Manager (9th edition)** exam: a study website and a spreadsheet of every question from the class materials, sorted into the book's 10 chapters.

## What's here

| File | What it is |
| --- | --- |
| `index.html` | The practice website. One file: download it and open it in any browser, on a phone or a computer. Nothing to install. |
| `servsafe-question-bank.xlsx` | Every question in one spreadsheet: question, all answer choices, correct answer, chapter (1–10), explanation and what it tests. |
| `questions.json` | The question bank both of the above are built from. |
| `build.py` | Rebuilds the spreadsheet and the website's data from `questions.json`. |

## The website

- **Mode 1 · Show answers right away:** every answer and explanation shows at once. Good for a first pass.
- **Mode 2 · Test me first:** answer each question, then see whether you were right, the explanation and what the question tests. You get a score at the end, a breakdown by chapter, and a button to retry the ones you missed. Questions you miss are remembered in your browser.
- **Mock exam:** 90 multiple-choice questions from all chapters with a 2-hour countdown, like the real exam. Answers and score at the end; 75% to pass.
- **Knowledge points:** the lists and short answers to know by heart (Big Six, FAT TOM, Big Nine allergens, HACCP principles, handwashing-station items and so on), with an option to hide the answers.
- **Timer:** optional stopwatch or countdown in the top bar.

All quiz questions are multiple choice. True/false, select-all, put-in-order and fill-in questions from the quizzes were rewritten as multiple choice; the spreadsheet's **Website rewrites** sheet shows each one next to its original.

## The spreadsheet

- **Questions:** 457 rows, one per question. Filter by chapter, source or topic with the arrows on row 1.
- **Yellow rows (43):** the source doesn't give the answer (for example a Mock Test question that was answered wrong, or a review slide that never states it), so an answer was assigned from ServSafe Manager 9th edition / FDA Food Code rules. The Explanation column gives the rule. These are also tagged "Assigned answer" on the website.
- **Pink rows (4):** review slides that depend on a picture that isn't in the file, so they can't be answered. They are kept in the spreadsheet only.
- **Summary**, **Assigned answers**, **Website rewrites** and **Topics** sheets.

## Sources

| Source | Questions |
| --- | --- |
| Kahoot Ch 4–6 practice questions | 24 (the export lists 27 items; 3 are chapter title slides) |
| SSM9E Activity Quizzes 1–10 | 112 |
| Practice Test | 40 |
| Diagnostic Test | 80 |
| Mock ServSafe Test (attempt review) | 50 |
| Review slide deck, all chapters | 151 |

The question content comes from ServSafe course materials © National Restaurant Association Educational Foundation (NRAEF), shared here for NUTR248 class study.

## Updating

Edit `questions.json`, then run:

```bash
python3 build.py
```

It rewrites `servsafe-question-bank.xlsx` and the question data inside `index.html`. It needs Python with `openpyxl`.
