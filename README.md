# Cold Email Generator

Generate HR- and company-specific cold emails from your resume and a company/HR spreadsheet.

## Folder layout

```text
main_input/
  resume/
    your_resume.pdf
  excel/
    company_hr_details.xlsx
output/
  generated_cold_emails.xlsx
```

## Spreadsheet columns

Your company/HR Excel or CSV file should include these columns:

```text
Company name | HR Name | HREMAIL
```

You can add more columns, such as `Industry`, `Location`, `Job Title`, `Role`, `Company Details`, or `Website`. The generator will pass those details into the email prompt.

## Setup

Install dependencies:

```powershell
pip install -r requirements.txt
```

LLM generation is optional. To use the Gemini API, create a `.env` file:

```powershell
Copy-Item .env.example .env
```

Then edit `.env` and set `GEMINI_API_KEY`.

The default model is `gemini-3-flash-preview`. If your Gemini account has a different Flash model ID, set `GEMINI_MODEL` in `.env` or pass `--model`.

## Run

Put your resume in `main_input/resume/` and your company/HR spreadsheet in `main_input/excel/`, then run:

```powershell
python cold_email_generator.py
```

The generated file will be:

```text
output/generated_cold_emails.xlsx
```

To force local template mode without using an LLM:

```powershell
python cold_email_generator.py --no-llm
```

To target different roles:

```powershell
python cold_email_generator.py --roles "Data Scientist, Data Engineer, Data Analyst, Business Analyst"
```

## Send generated emails

Add SMTP settings to `.env` first. For Gmail, create an app password and use that as
`SMTP_PASSWORD`.

Preview the existing generated Excel without sending:

```powershell
python cold_email_generator.py --send-existing-output --dry-run
```

Send from the existing `output/generated_cold_emails.xlsx`:

```powershell
python cold_email_generator.py --send-existing-output
```

Generate and send in one run:

```powershell
python cold_email_generator.py --send-emails
```

The sender uses `HREMAIL` as the recipient, derives `HR Name` from the mail ID,
uses the `Subject:` line from `Generated Email` as the email subject, and sends
the remaining text as the body. After a real send, the workbook is updated with
`Email Subject`, `Email Body`, `Email Send Status`, and `Email Sent At`.

Optional controls:

```powershell
python cold_email_generator.py --send-existing-output --send-limit 5
python cold_email_generator.py --send-existing-output --attach-resume
python cold_email_generator.py --send-existing-output --attachment path\to\file.pdf
```
