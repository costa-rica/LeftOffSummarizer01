"""
OneDrive file downloader using Microsoft Graph API.
Downloads the LEFT-OFF.docx file automatically using stored refresh token.
"""

import os
import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

# Load environment variables
load_dotenv()

APPLICATION_ID = os.getenv('APPLICATION_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
REFRESH_TOKEN = os.getenv('REFRESH_TOKEN')

# Microsoft Graph API configuration
# Note: 'offline_access' is automatically handled by MSAL and should not be included
SCOPES = ['Files.Read', 'Files.Read.All']

# Target file information (from README)
TARGET_FILE_ID = '60DF6D5FEED7AC60!s4170a85f60564e26947547219f91c294'
TARGET_FILE_NAME = 'LEFT-OFF.docx'


def get_access_token():
    """
    Gets a fresh access token using the stored refresh token.
    No browser interaction required.
    """
    if not REFRESH_TOKEN:
        print("ERROR: REFRESH_TOKEN not found in .env file")
        print("Please run get_auth_token.py first to obtain a refresh token.")
        return None

    # Create MSAL confidential client application (uses CLIENT_SECRET)
    app = ConfidentialClientApplication(
        APPLICATION_ID,
        client_credential=CLIENT_SECRET,
        authority='https://login.microsoftonline.com/consumers'
    )

    try:
        # Use refresh token to get new access token silently
        result = app.acquire_token_by_refresh_token(
            refresh_token=REFRESH_TOKEN,
            scopes=SCOPES
        )

        if 'access_token' in result:
            print("Access token obtained successfully (no browser needed)")

            # Update refresh token if a new one was issued
            if 'refresh_token' in result and result['refresh_token'] != REFRESH_TOKEN:
                print("\nNOTE: A new refresh token was issued.")
                print("Update your .env file with this new token:")
                print(f"REFRESH_TOKEN={result['refresh_token']}")
                print()

            return result['access_token']
        else:
            print("ERROR: Failed to get access token")
            print(f"Error: {result.get('error', 'Unknown error')}")
            print(f"Description: {result.get('error_description', '')}")

            if result.get('error') == 'invalid_grant':
                print("\nYour refresh token may have expired.")
                print("Please run get_auth_token.py again to get a new token.")

            return None

    except Exception as e:
        print(f"ERROR: Exception while getting access token: {e}")
        return None


def download_file(access_token, file_id, output_filename):
    """
    Downloads a file from OneDrive using Microsoft Graph API.

    Args:
        access_token: Valid access token for Microsoft Graph
        file_id: The ID of the file to download
        output_filename: Local filename to save the downloaded file
    """
    # Microsoft Graph API endpoint for file content
    download_url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content'

    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    print(f"\nDownloading file: {output_filename}")
    print(f"File ID: {file_id}")
    print("Making request to Microsoft Graph API...")

    try:
        response = requests.get(download_url, headers=headers, stream=True)

        if response.status_code == 200:
            # Save file to disk
            with open(output_filename, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            file_size = os.path.getsize(output_filename)
            print(f"\nSUCCESS! File downloaded successfully.")
            print(f"Saved to: {os.path.abspath(output_filename)}")
            print(f"File size: {file_size:,} bytes")
            return True
        else:
            print(f"\nERROR: Failed to download file")
            print(f"Status code: {response.status_code}")
            print(f"Response: {response.text}")
            return False

    except Exception as e:
        print(f"\nERROR: Exception while downloading file: {e}")
        return False


def main():
    """Main application logic"""
    print("=" * 70)
    print("ONEDRIVE FILE DOWNLOADER")
    print("=" * 70)
    print()

    # Validate environment variables
    if not APPLICATION_ID:
        print("ERROR: APPLICATION_ID not found in .env file")
        return

    # Step 1: Get access token using refresh token
    print("Step 1: Obtaining access token...")
    access_token = get_access_token()

    if not access_token:
        print("\nFailed to obtain access token. Exiting.")
        return

    # Step 2: Download the target file
    print("\nStep 2: Downloading file from OneDrive...")
    success = download_file(access_token, TARGET_FILE_ID, TARGET_FILE_NAME)

    if success:
        print("\n" + "=" * 70)
        print("DOWNLOAD COMPLETE")
        print("=" * 70)
    else:
        print("\nDownload failed. Please check the error messages above.")


if __name__ == '__main__':
    main()
