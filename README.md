# LeftOffSummarizer

A Python service that automatically downloads your LEFT-OFF.docx file from OneDrive, extracts the last 7 days of activities, and generates an AI-powered summary using OpenAI.

## Features

- Downloads LEFT-OFF.docx from OneDrive using Microsoft Graph API
- Parses the document to extract the most recent 7 days of activities
- Generates intelligent summaries using OpenAI's gpt-4o-mini model
- Comprehensive logging for tracking progress and debugging

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment Variables

Copy `.env.example` to `.env` and fill in your credentials:

```bash
cp .env.example .env
```

Edit `.env` with your actual values:
- `PATH_LEFT_OFF_SUMMARIZER`: Directory where output files will be saved
- `TARGET_FILE_ID`: OneDrive file ID for LEFT-OFF.docx
- `APPLICATION_ID`, `CLIENT_SECRET`, `REFRESH_TOKEN`: Microsoft Graph API credentials
- `KEY_OPENAI`: Your OpenAI API key
- `URL_BASE_OPENAI`: OpenAI API base URL (usually https://api.openai.com/v1)

### 3. Create Prompt File

Create a `prompt.md` file in your `PATH_LEFT_OFF_SUMMARIZER` directory. This file should contain your summarization instructions. Use the placeholder `<< last-7-days-activities.md >>` where you want the activities content to be inserted.

Example `prompt.md`:
```markdown
Please summarize the following activities from the last 7 days:

<< last-7-days-activities.md >>

Provide a concise summary highlighting key accomplishments and ongoing tasks.
```

## Usage

Run the service:

```bash
python src/main.py
```

The service will:
1. Download LEFT-OFF.docx from OneDrive
2. Extract activities from the last 7 days
3. Generate a summary using OpenAI

### Time Window Restriction

For safety, the service only runs during its scheduled window: **Sunday 10:55 PM - 11:05 PM** (local system time). This prevents accidental executions.

To bypass this restriction for manual runs or testing:

```bash
python src/main.py --run-anyway
```

## Output Files

All files are saved to the directory specified in `PATH_LEFT_OFF_SUMMARIZER`:

- `LEFT-OFF.docx`: Downloaded document from OneDrive
- `last-7-days-activities.md`: Extracted activities from the last 7 days
- `last-7-days-activities-summary.md`: AI-generated summary

## Document Format

The LEFT-OFF.docx file should be structured as follows:
- Heading 1: Date in YYYYMMDD format (e.g., 20231115)
- Heading 2: "LEFT-OFF" and "Accomplished Today" sections
- Most recent entries at the top, oldest at the bottom

## Logging

The service uses Python's logging module to track progress. All logs are output to the console with timestamps and severity levels.

## Troubleshooting

- **Authentication errors**: Ensure your REFRESH_TOKEN is valid. You may need to regenerate it.
- **File not found**: Verify the TARGET_FILE_ID points to the correct OneDrive file.
- **OpenAI errors**: Check that KEY_OPENAI is valid and you have sufficient API credits.
