# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

LeftOffSummarizer is a Python service that automates the process of downloading a LEFT-OFF.docx file from OneDrive, extracting the most recent 7 days of activities, and generating an AI-powered summary using OpenAI.

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the main service (only allowed on Sunday 10:55 PM - 11:05 PM)
python src/main.py

# Run the service anytime (bypass time window restriction)
python src/main.py --run-anyway
```

### Time Window Guardrail

The service includes a time window restriction to prevent accidental execution outside of the scheduled cron job window:
- **Allowed window**: Sunday 10:55 PM - 11:05 PM (local system time)
- **Buffer**: 5 minutes before and after 11:00 PM for clock synchronization tolerance
- **Exit behavior**: If run outside this window without the `--run-anyway` flag, the service logs a warning and exits with code 2
- **Bypass**: Use `--run-anyway` flag to run the service at any time for testing or manual execution

## Architecture

The application follows a three-stage pipeline architecture:

1. **OneDrive Download** (`src/onedrive_client.py`)
   - Uses Microsoft Graph API with MSAL (Microsoft Authentication Library)
   - Authenticates via refresh token to obtain access tokens
   - Downloads the target LEFT-OFF.docx file from OneDrive

2. **Document Parsing** (`src/document_parser.py`)
   - Parses .docx files using python-docx library
   - Extracts content from the last 7 days based on Heading 1 dates in YYYYMMDD format
   - Converts document structure to markdown format
   - Document structure expectation: Most recent entries at top, oldest at bottom

3. **AI Summary Generation** (`src/summary_generator.py`)
   - Uses OpenAI API (gpt-4o-mini by default)
   - Reads a customizable prompt template from `prompt.md` in the output directory
   - Replaces the placeholder `<< last-7-days-activities.md >>` with extracted activities
   - Prefixes the generated summary with today's date (YYYY-MM-DD)

The orchestration happens in `src/main.py` which:
- Loads environment variables using python-dotenv
- Executes the three stages sequentially with comprehensive logging
- Returns exit code 0 on success, 1 on failure

## Configuration

All configuration is managed via environment variables in `.env`:

Required variables:
- `PATH_LEFT_OFF_SUMMARIZER`: Output directory for all files
- `TARGET_FILE_ID`: OneDrive file ID for LEFT-OFF.docx
- `APPLICATION_ID`, `CLIENT_SECRET`, `REFRESH_TOKEN`: Microsoft Graph API credentials
- `KEY_OPENAI`: OpenAI API key
- `URL_BASE_OPENAI`: OpenAI API base URL

A `prompt.md` file must exist in `PATH_LEFT_OFF_SUMMARIZER` containing the summarization instructions with the placeholder `<< last-7-days-activities.md >>`.

## Output Files

All output files are saved to `PATH_LEFT_OFF_SUMMARIZER`:
- `LEFT-OFF.docx`: Downloaded document
- `last-7-days-activities.md`: Extracted last 7 days
- `last-7-days-activities-summary.md`: AI-generated summary (prefixed with YYYY-MM-DD date)

## Document Format Requirements

The LEFT-OFF.docx file must follow this structure:
- **Heading 1**: Date in YYYYMMDD format (e.g., 20231115)
- **Heading 2**: Section headers like "LEFT-OFF" and "Accomplished Today"
- **Organization**: Most recent entries at the top, oldest at the bottom

The parser extracts all content from the beginning of the document until it finds the first Heading 1 date that is 8+ days old (the cutoff ensures 7 full days of content).

## Error Handling

The application uses Python's logging module throughout. All modules have their own logger instances. Common failure points:
- Invalid/expired refresh tokens (OneDrive authentication)
- Missing or incorrectly formatted document structure
- Missing prompt.md file
- OpenAI API errors (invalid key, rate limits, insufficient credits)

### Exit Codes

- **0**: Success - all operations completed successfully
- **1**: Error - operational failure (auth error, file error, API error, etc.)
- **2**: Time window restriction - execution attempted outside allowed Sunday 10:55 PM - 11:05 PM window
