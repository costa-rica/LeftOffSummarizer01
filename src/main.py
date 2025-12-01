"""
LeftOffSummarizer - Main entry point

This service:
1. Downloads LEFT-OFF.docx from OneDrive
2. Extracts the last 7 days of activities
3. Generates a summary using OpenAI
"""

import os
import sys
import logging
import argparse
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

from onedrive_client import OneDriveClient
from document_parser import DocumentParser
from summary_generator import SummaryGenerator


def setup_logging():
    """Configure logging for the application."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )


def is_within_allowed_window():
    """
    Check if current time is within the allowed execution window.

    Allowed window: Sunday 10:55 PM - 11:05 PM (local system time)

    Returns:
        bool: True if within allowed window, False otherwise
    """
    now = datetime.now()

    # Check if it's Sunday (weekday 6)
    if now.weekday() != 6:
        return False

    # Check if time is between 22:55 (10:55 PM) and 23:05 (11:05 PM)
    current_time = now.time()
    start_time = datetime.strptime("22:55", "%H:%M").time()
    end_time = datetime.strptime("23:05", "%H:%M").time()

    return start_time <= current_time <= end_time


def parse_arguments():
    """
    Parse command-line arguments.

    Returns:
        argparse.Namespace: Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description='LeftOffSummarizer - Summarize recent LEFT-OFF activities'
    )
    parser.add_argument(
        '--run-anyway',
        action='store_true',
        help='Run the service regardless of the scheduled time window'
    )
    return parser.parse_args()


def main():
    """Main application logic."""
    setup_logging()
    logger = logging.getLogger(__name__)

    # Parse command-line arguments
    args = parse_arguments()

    # Check if we're within the allowed time window
    if not args.run_anyway and not is_within_allowed_window():
        now = datetime.now()
        logger.warning("=" * 70)
        logger.warning("EXECUTION BLOCKED")
        logger.warning("=" * 70)
        logger.warning(f"Current time: {now.strftime('%A, %Y-%m-%d %H:%M:%S')}")
        logger.warning("Service can only run on Sunday between 10:55 PM and 11:05 PM")
        logger.warning("Use --run-anyway flag to bypass this restriction")
        logger.warning("=" * 70)
        return 2  # Exit code 2 for "outside allowed window"

    logger.info("=" * 70)
    logger.info("LEFT-OFF SUMMARIZER")
    logger.info("=" * 70)

    if args.run_anyway:
        logger.info("Running with --run-anyway flag (bypassing time window check)")

    # Load environment variables
    load_dotenv()

    # Get configuration from environment
    app_name = os.getenv('NAME_APP', 'LeftOffSummarizer')
    base_path = os.getenv('PATH_LEFT_OFF_SUMMARIZER')
    target_file_name = os.getenv('NAME_TARGET_FILE', 'LEFT-OFF.docx')
    target_file_id = os.getenv('TARGET_FILE_ID')
    application_id = os.getenv('APPLICATION_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    refresh_token = os.getenv('REFRESH_TOKEN')
    openai_key = os.getenv('KEY_OPENAI')
    openai_base_url = os.getenv('URL_BASE_OPENAI')

    # Validate required environment variables
    required_vars = {
        'PATH_LEFT_OFF_SUMMARIZER': base_path,
        'TARGET_FILE_ID': target_file_id,
        'APPLICATION_ID': application_id,
        'CLIENT_SECRET': client_secret,
        'REFRESH_TOKEN': refresh_token,
        'KEY_OPENAI': openai_key,
        'URL_BASE_OPENAI': openai_base_url
    }

    missing_vars = [var for var, value in required_vars.items() if not value]
    if missing_vars:
        logger.error(f"Missing required environment variables: {', '.join(missing_vars)}")
        return 1

    # Set up file paths
    docx_path = os.path.join(base_path, target_file_name)
    activities_path = os.path.join(base_path, 'last-7-days-activities.md')
    summary_path = os.path.join(base_path, 'last-7-days-activities-summary.md')
    prompt_path = os.path.join(base_path, 'prompt.md')

    try:
        # Step 1: Download file from OneDrive
        logger.info("\n" + "=" * 70)
        logger.info("STEP 1: Downloading file from OneDrive")
        logger.info("=" * 70)

        onedrive = OneDriveClient(application_id, client_secret, refresh_token)

        if not onedrive.get_access_token():
            logger.error("Failed to obtain access token")
            return 1

        if not onedrive.download_file(target_file_id, docx_path):
            logger.error("Failed to download file")
            return 1

        # Step 2: Parse document and extract last 7 days
        logger.info("\n" + "=" * 70)
        logger.info("STEP 2: Extracting last 7 days of activities")
        logger.info("=" * 70)

        parser = DocumentParser(docx_path)

        if not parser.load_document():
            logger.error("Failed to load document")
            return 1

        if not parser.extract_last_7_days(activities_path):
            logger.error("Failed to extract activities")
            return 1

        # Step 3: Generate summary using OpenAI
        logger.info("\n" + "=" * 70)
        logger.info("STEP 3: Generating summary with OpenAI")
        logger.info("=" * 70)

        generator = SummaryGenerator(openai_key, openai_base_url)

        if not generator.generate_summary(prompt_path, activities_path, summary_path):
            logger.error("Failed to generate summary")
            return 1

        # Success
        logger.info("\n" + "=" * 70)
        logger.info("SUCCESS - All steps completed")
        logger.info("=" * 70)
        logger.info(f"Activities extracted to: {activities_path}")
        logger.info(f"Summary saved to: {summary_path}")

        return 0

    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return 1


if __name__ == '__main__':
    sys.exit(main())
