import logging
import os
import re
from azure.identity import ClientSecretCredential
import requests

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url):
        """Initialize SharePoint client without immediate authentication"""
        self.site_url = site_url.rstrip('/')  # Remove trailing slash
        self.ctx = None
        self.tenant = self._extract_tenant_from_url(site_url)
        self.access_token = None
        logger.info(f"Initialized SharePoint client for site: {site_url}")

    def _extract_tenant_from_url(self, url):
        """Extract tenant name from SharePoint URL"""
        match = re.search(r'https://([^.]+)\.sharepoint\.com', url)
        if match:
            return match.group(1)
        raise ValueError("Invalid SharePoint URL format. Expected: https://<tenant>.sharepoint.com/...")

    def authenticate(self):
        """Initialize authentication using Azure AD credentials"""
        try:
            logger.info("Starting SharePoint authentication process...")

            # Get Azure AD app credentials
            client_id = os.environ.get('AZURE_CLIENT_ID')
            client_secret = os.environ.get('AZURE_CLIENT_SECRET')
            tenant_id = os.environ.get('AZURE_TENANT_ID')

            if not all([client_id, client_secret, tenant_id]):
                missing = []
                if not client_id: missing.append("AZURE_CLIENT_ID")
                if not client_secret: missing.append("AZURE_CLIENT_SECRET")
                if not tenant_id: missing.append("AZURE_TENANT_ID")
                raise ValueError(f"Missing required credentials: {', '.join(missing)}")

            logger.info(f"Authenticating with SharePoint site: {self.site_url}")
            logger.info(f"Using tenant ID: {tenant_id}")

            # Create credential object
            credential = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret
            )

            # Request token with Microsoft Graph API scope
            scopes = ["https://graph.microsoft.com/.default"]
            logger.info(f"Requesting token for scopes: {scopes}")

            try:
                token_response = credential.get_token(scopes[0])
                self.access_token = token_response.token
                logger.info("Token acquired successfully")
            except Exception as token_error:
                logger.error(f"Failed to acquire token: {str(token_error)}")
                raise

            # Test connection using Microsoft Graph API
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }

            # Use Graph API to verify site access
            site_id = f"sites/{self.tenant}.sharepoint.com:/"
            test_url = f"https://graph.microsoft.com/v1.0/{site_id}"
            logger.info(f"Testing SharePoint connection via Graph API at: {test_url}")

            try:
                response = requests.get(test_url, headers=headers)
                logger.info(f"Response status code: {response.status_code}")

                if response.status_code == 200:
                    logger.info("Successfully authenticated with SharePoint")
                    return True
                else:
                    error_content = response.text
                    logger.error(f"Graph API Response Status: {response.status_code}")
                    logger.error(f"Graph API Response Content: {error_content}")

                    if response.status_code == 401:
                        logger.error("Authentication failed - invalid credentials or insufficient permissions")
                        logger.error("Please verify:")
                        logger.error("1. The Azure AD application has the Microsoft Graph API permissions")
                        logger.error("2. Admin consent is granted for Sites.Read.All and Sites.ReadWrite.All")
                        logger.error("3. The application is properly configured for client credentials flow")

                    raise Exception(f"Failed to connect to SharePoint via Graph API. Status code: {response.status_code}")

            except requests.exceptions.RequestException as req_error:
                logger.error(f"Request failed: {str(req_error)}")
                raise

        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            logger.error("Please ensure:")
            logger.error("1. The Azure AD application is properly registered")
            logger.error("2. Admin consent is granted for Microsoft Graph API permissions")
            logger.error("3. The Client ID and Secret are correct")
            logger.error("4. The SharePoint site URL is correct")
            logger.error("5. The Tenant ID matches your Azure AD tenant")
            raise

    def get_libraries(self):
        """Get all document libraries in the site using Microsoft Graph API"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }

            site_id = f"sites/{self.tenant}.sharepoint.com:/"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers=headers)

            if response.status_code == 200:
                data = response.json()
                return [drive['name'] for drive in data.get('value', [])]
            else:
                logger.error(f"Failed to get libraries. Response: {response.text}")
                raise Exception(f"Failed to get libraries. Status code: {response.status_code}")

        except Exception as e:
            logger.error(f"Failed to get libraries: {str(e)}")
            raise

    def get_files(self, library_name):
        """Get all files in a library using Microsoft Graph API"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }

            # First get the drive ID for the library
            site_id = f"sites/{self.tenant}.sharepoint.com:/"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers=headers)

            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            drives_data = response.json()
            drive_id = None
            for drive in drives_data.get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found")

            # Get files from the drive
            files_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
            response = requests.get(files_url, headers=headers)

            if response.status_code == 200:
                data = response.json()
                files = []
                for item in data.get('value', []):
                    if 'file' in item:
                        files.append({
                            'Id': item['id'],
                            'Name': item['name'],
                            'Path': item['webUrl']
                        })
                return files
            else:
                logger.error(f"Failed to get files. Response: {response.text}")
                raise Exception(f"Failed to get files. Status code: {response.status_code}")

        except Exception as e:
            logger.error(f"Failed to get files from library {library_name}: {str(e)}")
            raise

    def rename_file(self, library_name, old_name, new_name):
        """Rename a file using Microsoft Graph API"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            }

            # First get the drive ID for the library
            site_id = f"sites/{self.tenant}.sharepoint.com:/"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers=headers)

            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            drives_data = response.json()
            drive_id = None
            for drive in drives_data.get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found")

            # Search for the file
            search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='{old_name}')"
            response = requests.get(search_url, headers=headers)

            if response.status_code != 200:
                raise Exception(f"Failed to find file. Status code: {response.status_code}")

            items = response.json().get('value', [])
            file_id = None
            for item in items:
                if item.get('name') == old_name:
                    file_id = item['id']
                    break

            if not file_id:
                raise Exception(f"File '{old_name}' not found")

            # Rename the file
            update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
            update_data = {'name': new_name}

            response = requests.patch(update_url, headers=headers, json=update_data)

            if response.status_code not in [200, 201]:
                logger.error(f"Failed to rename file. Response: {response.text}")
                raise Exception(f"Failed to rename file. Status code: {response.status_code}")

            logger.info(f"File renamed successfully from {old_name} to {new_name}")

        except Exception as e:
            logger.error(f"Failed to rename file {old_name}: {str(e)}")
            raise