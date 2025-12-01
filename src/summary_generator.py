"""
Summary generator using OpenAI API.
"""

import logging
from openai import OpenAI


logger = logging.getLogger(__name__)


class SummaryGenerator:
    """Generates summaries using OpenAI API."""

    def __init__(self, api_key, base_url):
        """
        Initialize summary generator.

        Args:
            api_key: OpenAI API key
            base_url: OpenAI base URL
        """
        self.client = OpenAI(api_key=api_key, base_url=base_url)

    def generate_summary(self, prompt_path, activities_path, output_path, model='gpt-4o-mini'):
        """
        Generates a summary of activities using OpenAI.

        Args:
            prompt_path: Path to the prompt.md file
            activities_path: Path to the last-7-days-activities.md file
            output_path: Path to save the summary
            model: OpenAI model to use (default: gpt-4o-mini)

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Read the prompt template
            logger.info(f"Reading prompt from: {prompt_path}")
            with open(prompt_path, 'r', encoding='utf-8') as f:
                prompt_template = f.read()

            # Read the activities content
            logger.info(f"Reading activities from: {activities_path}")
            with open(activities_path, 'r', encoding='utf-8') as f:
                activities_content = f.read()

            # Replace placeholder with activities content
            prompt = prompt_template.replace('<< last-7-days-activities.md >>', activities_content)

            # Call OpenAI API
            logger.info(f"Generating summary using {model}")
            response = self.client.chat.completions.create(
                model=model,
                messages=[
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ]
            )

            # Extract summary from response
            summary = response.choices[0].message.content
            logger.info(f"Summary generated ({len(summary)} characters)")

            # Save summary to file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(summary)

            logger.info(f"Summary saved to: {output_path}")
            return True

        except FileNotFoundError as e:
            logger.error(f"File not found: {e}")
            return False
        except Exception as e:
            logger.error(f"Failed to generate summary: {e}")
            return False
