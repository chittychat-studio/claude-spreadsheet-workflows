# Every prompt from the video

Copy a prompt, attach the files it names (from [`sample-files/`](sample-files/)), then check the result before you trust it.

## Merging lists

### 1. Look before you change

```text
I've uploaded members_existing.xlsx. Before you change anything, tell me what's in it: the columns, how many rows, and anything that looks inconsistent. Don't edit the file yet.
```

**Attach:** members_existing.xlsx  
**Check afterwards:** It describes the file without editing it. Columns, row count and inconsistencies match what you see.

### 2. Merge two lists

```text
I've attached two files: members_existing.xlsx and new_submissions.xlsx. Merge them into one member list in a new Excel file. Treat two rows as the same person only if the email matches exactly. If names look similar but the emails are different, do not merge them: put them on a separate sheet called Review, with the reason. Keep phone numbers as text so leading zeros stay. Add a Summary sheet showing how many rows came from each file, how many were merged, and how many need review. Do not delete anything.
```

**Attach:** members_existing.xlsx, new_submissions.xlsx  
**Check afterwards:** Exact-email matches merged; look-alike names on the Review sheet with a reason; phones keep the leading 0; nothing deleted.

### 3. Ask for the evidence

```text
Show me exactly which rows you put on the Review sheet, and why.
```

**Attach:** (same chat)  
**Check afterwards:** Each Review row has a reason you agree with. You make the final call on every possible match.

### 4. Add a status + print copy

```text
Using the merged member list you just made, add a Status column. Use Renewed if the person appears in both files with the same email, Not renewed if they are only in members_existing.xlsx, and New if they are only in new_submissions.xlsx. Sort by Status, then by Full Name. Then give me a second copy of the same list, set up for printing: large font, the table fits the page width, and nothing depends on colour.
```

**Attach:** (same chat)  
**Check afterwards:** Renewed / Not renewed / New follow the rule; print copy fits the page and does not rely on colour.

## Paper forms

### 5. Photo to table

```text
This is a photo of a paper sign-up sheet. Turn it into a table with these columns: Name, Phone, Email. Copy exactly what is written. Keep phone numbers as text, including any leading zero. If you can't read something clearly, write UNCLEAR in that cell. Do not guess. At the end, tell me how many rows you found.
```

**Attach:** signup_sheet_photo.png  
**Check afterwards:** Anything unreadable says UNCLEAR, never a guess. Spot-check names and phone numbers against the photo.

### 6. Get the Excel file

```text
Give me that table as an Excel file, with the Phone column stored as text.
```

**Attach:** (same chat)  
**Check afterwards:** Phone column stored as text, so leading zeros survive.

## Inherited workbook

### 7. Audit before you edit

```text
I've inherited this workbook and I need to understand it before I change anything. For each sheet, explain what it does in plain English. List the input cells, the calculated cells, and how the sheets connect. Flag anything risky: numbers typed in where I'd expect a formula, formulas that look inconsistent, or errors. Also list the cells I should never overwrite. Do not change the file. Just explain.
```

**Attach:** inherited_budget_workbook.xlsx  
**Check afterwards:** It flags typed numbers where a formula should be, and lists cells never to overwrite. The file is unchanged.

### 8. Explain one formula

```text
Explain this formula like I'm new to spreadsheets. Step by step. No jargon.
```

**Attach:** (click the formula's cell first)  
**Check afterwards:** You can explain the formula back in your own words.

## Rotas

### 9. Build a rota with rules

```text
Using rota_volunteers.xlsx, build a rota for next month as a new Excel file. Rules: never schedule anyone on a date they've marked unavailable. Nobody works two shifts in a row. Spread shifts as evenly as possible. Put the calendar on one sheet, and on a second sheet add a summary showing each person's total shifts. List any rule you couldn't meet, and why.
```

**Attach:** rota_volunteers.xlsx  
**Check afterwards:** Check the calendar against every rule yourself. Equal totals do not prove the rota is valid.

### 10. Draft the email macro

```text
Write an Excel VBA macro called SendRotaEmail. It should attach the workbook I have open, address an email to coordinator@example.org, use the subject "Volunteer rota for next month", and add a short, friendly message from Ravi. Use desktop Outlook. Show the email as a draft with .Display. Never use .Send. Save the workbook before attaching it. Add a short comment above each step.
```

**Attach:** (your rota workbook)  
**Check afterwards:** Read the code before running it. It must use .Display (a draft you send yourself), never .Send. Trusted files only.

## Power-user moves

### 11. Calculate with code

```text
Using event_attendance.xlsx, calculate the total and the average attendance for each event type using code. Show the results in a table and explain how you calculated them.
```

**Attach:** event_attendance.xlsx  
**Check afterwards:** A blank month is not a zero. Check one total by hand.

### 12. Real formulas, not typed numbers

```text
Add a totals row to this sheet using real Excel formulas, like SUM, not typed numbers, so the totals update when I change a value. Give me the updated Excel file.
```

**Attach:** (any sheet with numbers)  
**Check afterwards:** Change one value: the totals update. If they don't, they were typed in.

### 13. Fix dates and phone numbers

```text
In new_submissions.xlsx, the Join date column uses mixed date formats. Convert every date to day, month, year format, written as DD/MM/YYYY. If a date is ambiguous, like 03/04, do not guess: flag it in a new column called Check. Also make sure the Phone column is stored as text so leading zeros are kept. Give me the corrected Excel file.
```

**Attach:** new_submissions.xlsx  
**Check afterwards:** Ambiguous dates are flagged in Check, not guessed. Phones keep the leading 0.

### 14. Write a formula with you

```text
In new_submissions.xlsx, the Join date column holds the date each person joined. I'm in the UK, so dates are day, month, year. Write me an Excel formula for a new column that says Over 30 days if the person joined more than 30 days before today, and Recent if not. Explain the formula in one sentence, and tell me which rows it can't handle.
```

**Attach:** new_submissions.xlsx  
**Check afterwards:** Check one clean row by hand. Note the rows it says it can't handle.

## Insights and charts

### 15. Summary + chart

```text
Using event_attendance.xlsx, give me a short summary of attendance trends for each event type. Then create a line chart of monthly attendance with one line per event type. Label both axes, add a title, and note the data source under the chart. Only use numbers from the file.
```

**Attach:** event_attendance.xlsx  
**Check afterwards:** Every number in the summary exists in the file. Labels, title and source are on the chart.

### 16. Check the chart against the file

```text
I'm checking the chart against the file. For Jun 2026, show me the three values exactly as they appear in event_attendance.xlsx, next to the values the chart plots for that month. Tell me if anything differs.
```

**Attach:** (same chat)  
**Check afterwards:** The values match the table exactly. A chart is a claim: check it.

## Everyday workflows

### 17. Meeting notes to action table

```text
Here are my meeting notes, attached as meeting_notes.txt. Pull out every action item into a table with these columns: Task, Owner, Due date, Source line. Only use what is written in the notes. If an owner or a date is not stated, write NOT STATED. Do not fill gaps with guesses.
```

**Attach:** meeting_notes.txt  
**Check afterwards:** Missing owners and dates say NOT STATED. Nothing is invented.

## Features (Projects)

### 18. Ask about project knowledge

```text
Using the member list in this project's knowledge, how many members are marked Active, and which column shows that?
```

**Attach:** members_existing.xlsx (added to a Project)  
**Check afterwards:** The answer names the column it used. Check the count yourself.

### 19. Standing project instructions

```text
Always keep phone numbers as text. Always use day, month, year for dates.
```

**Attach:** (Project instructions box)  
**Check afterwards:** New chats in the project follow the rules without being told again.

## Before you trust any result

1. Understand the source data
2. Describe the desired outcome
3. Request the appropriate processing
4. Inspect the generated result
5. Verify calculations and ambiguous records
6. Preserve important formatting
7. Review privacy and data-handling considerations
8. Save the final output clearly

---

> AI-generated data, formulas, and spreadsheet modifications can contain errors or hallucinated numbers that look completely correct. Always verify the results with your own eyes before using them for operational decisions, reports, or financial calculations.
