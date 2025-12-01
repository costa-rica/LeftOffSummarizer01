"""
OneDrive client for downloading files using Microsoft Graph API.
"""

import os
import logging
import requests
from msal import ConfidentialClientApplication


logger = logging.getLogger(__name__)


class OneDriveClient:
    """Client for interacting with OneDrive via Microsoft Graph API."""

    SCOPES = ['Files.Read', 'Files.Read.All']
    AUTHORITY = 'https://login.microsoftonline.com/consumers'
    GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

    def __init__(self, application_id, client_secret, refresh_token):
        """
        Initialize OneDrive client.

        Args:
            application_id: Azure AD application ID
            client_secret: Azure AD client secret
            refresh_token: Refresh token for authentication
        """
        self.application_id = application_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self.access_token = None

    def get_access_token(self):
        """
        Gets a fresh access token using the stored refresh token.

        Returns:
            str: Access token if successful, None otherwise
        """
        logger.info("Obtaining access token from Microsoft Graph API")

        app = ConfidentialClientApplication(
            self.application_id,
            client_credential=self.client_secret,
            authority=self.AUTHORITY
        )

        try:
            result = app.acquire_token_by_refresh_token(
                refresh_token=self.refresh_token,
                scopes=self.SCOPES
            )

            if 'access_token' in result:
                logger.info("Access token obtained successfully")

                # Check if new refresh token was issued
                if 'refresh_token' in result and result['refresh_token'] != self.refresh_token:
                    logger.warning("New refresh token issued - update your .env file")
                    logger.warning(f"REFRESH_TOKEN={result['refresh_token']}")

                self.access_token = result['access_token']
                return self.access_token
            else:
                error = result.get('error', 'Unknown error')
                error_desc = result.get('error_description', '')
                logger.error(f"Failed to get access token: {error} - {error_desc}")

                if error == 'invalid_grant':
                    logger.error("Refresh token may have expired")

                return None

        except Exception as e:
            logger.error(f"Exception while getting access token: {e}")
            return None

    def download_file(self, file_id, output_path):
        """
        Downloads a file from OneDrive.

        Args:
            file_id: The ID of the file to download
            output_path: Local path to save the downloaded file

        Returns:
            bool: True if successful, False otherwise
        """
        if not self.access_token:
            logger.error("No access token available")
            return False

        download_url = f'{self.GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/content'
        headers = {'Authorization': f'Bearer {self.access_token}'}

        logger.info(f"Downloading file from OneDrive (ID: {file_id})")
        logger.info(f"Saving to: {output_path}")

        try:
            response = requests.get(download_url, headers=headers, stream=True)

            if response.status_code == 200:
                # Ensure output directory exists
                os.makedirs(os.path.dirname(output_path), exist_ok=True)

                with open(output_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

                file_size = os.path.getsize(output_path)
                logger.info(f"File downloaded successfully ({file_size:,} bytes)")
                return True
            else:
                logger.error(f"Failed to download file: HTTP {response.status_code}")
                logger.error(f"Response: {response.text}")
                return False

        except Exception as e:
            logger.error(f"Exception while downloading file: {e}")
            return False
