# Claude for Everyday Spreadsheet Workflows: prompt pack

Every prompt from the video, plus the sample files used on screen, so you can follow along step by step.

**Video:** Claude for Everyday Spreadsheet Workflows, on [Corporate Chitty Chat](https://www.youtube.com/@Corporatechittychat) (YouTube)

## What's here

| File | What it is |
|---|---|
| [`Claude_Spreadsheet_Prompt_Pack.xlsx`](Claude_Spreadsheet_Prompt_Pack.xlsx) | The full pack: all 19 prompts with the files they need and what to check afterwards, practice tasks, the 8-step checklist, an answer key for the sample files, and the email macro |
| [`prompts.md`](prompts.md) | The same prompts as a page you can read and copy right here on GitHub |
| [`sample-files/`](sample-files/) | The exact practice files from the video (all fictional data) |
| [`macro/SendRotaEmail.bas`](macro/SendRotaEmail.bas) | The rota email macro from the video. It opens a draft; it never sends |

## The method

1. **Describe** what you have and what you want.
2. **Process**: let Claude do the heavy lifting.
3. **Review** what comes back.
4. **Verify** the numbers that matter, with your own eyes.

## How to use it

1. Download the files in `sample-files/` (or the whole repo with **Code > Download ZIP**).
2. Open a chat with Claude, paste a prompt from `prompts.md` or the spreadsheet, and attach the files it names.
3. Before trusting the result, use the **Check afterwards** note for that prompt.
4. Then try it on your own files, after checking your company's rules on AI tools.

## About the sample data

Every name, email and phone number is invented. Emails use `example.com` / `example.org`, which are reserved for examples, and phone numbers use the `07700 900xxx` range Ofcom reserves for fiction. The files contain deliberate traps (duplicates, lost leading zeros, ambiguous dates, a typed total that doesn't add up). The **Answer key** tab lists them, so try the prompts before you peek.

## About the macro

`SendRotaEmail` needs desktop Excel on Windows and the classic desktop Outlook app. It saves the workbook, attaches it, and opens the email as a **draft** (`.Display`). You read it and press Send yourself. Only run macros in files you trust, read the code first, and change the recipient address before you run it.

## Important

AI-generated data, formulas, and spreadsheet modifications can contain errors or hallucinated numbers that look completely correct. Always verify the results with your own eyes before using them for operational decisions, reports, or financial calculations.

Some features shown in the video (Projects, the Excel add-in, connectors) depend on your plan. Check the official Claude help centre for what's available to you.

*Not affiliated with or endorsed by Anthropic. Claude is a product of Anthropic.*
