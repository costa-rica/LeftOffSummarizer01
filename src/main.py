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


def main():
    """Main application logic."""
    setup_logging()
    logger = logging.getLogger(__name__)

    logger.info("=" * 70)
    logger.info("LEFT-OFF SUMMARIZER")
    logger.info("=" * 70)

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
