import logging
import os
import re
from azure.identity import ClientSecretCredential
import requests

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url):
        """Initialize SharePoint client without immediate authentication"""
        self.site_url = site_url
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
            tenant_id = f"{self.tenant}.onmicrosoft.com"

            if not client_id or not client_secret:
                raise ValueError("AZURE_CLIENT_ID and AZURE_CLIENT_SECRET must be set")

            logger.info(f"Authenticating with SharePoint site: {self.site_url}")

            # Create credential object
            credential = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret
            )

            # Get access token with specific SharePoint scopes
            scopes = [
                f"https://{self.tenant}.sharepoint.com/Sites.Read.All",
                f"https://{self.tenant}.sharepoint.com/Sites.ReadWrite.All"
            ]
            self.access_token = credential.get_token(scopes[0]).token #Using the first scope for now.  Consider handling both.

            # Test connection
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json;odata=verbose'
            }
            response = requests.get(f"{self.site_url}/_api/web", headers=headers)

            if response.status_code == 200:
                logger.info("Successfully authenticated with SharePoint")
                return True
            else:
                logger.error(f"Response content: {response.text}")
                raise Exception(f"Failed to connect to SharePoint. Status code: {response.status_code}")

        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            logger.error("Please ensure:")
            logger.error("1. The Azure AD application is properly registered")
            logger.error("2. Admin consent is granted for SharePoint API permissions")
            logger.error("3. The Client ID and Secret are correct")
            logger.error("4. The SharePoint site URL is correct")
            raise

    def get_libraries(self):
        """Get all document libraries in the site"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json;odata=verbose'
            }

            response = requests.get(
                f"{self.site_url}/_api/web/lists?$filter=BaseTemplate eq 101",
                headers=headers
            )

            if response.status_code == 200:
                data = response.json()
                return [lib['Title'] for lib in data['d']['results']]
            else:
                raise Exception(f"Failed to get libraries. Status code: {response.status_code}")

        except Exception as e:
            logger.error(f"Failed to get libraries: {str(e)}")
            raise

    def get_files(self, library_name):
        """Get all files in a library"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json;odata=verbose'
            }

            response = requests.get(
                f"{self.site_url}/_api/web/lists/GetByTitle('{library_name}')/Items",
                headers=headers
            )

            if response.status_code == 200:
                data = response.json()
                files = []
                for item in data['d']['results']:
                    if 'FileLeafRef' in item:
                        files.append({
                            'Id': item['ID'],
                            'Name': item['FileLeafRef'],
                            'Path': item['FileRef']
                        })
                return files
            else:
                raise Exception(f"Failed to get files. Status code: {response.status_code}")

        except Exception as e:
            logger.error(f"Failed to get files from library {library_name}: {str(e)}")
            raise

    def rename_file(self, library_name, old_name, new_name):
        """Rename a file in SharePoint"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-RequestDigest': self._get_form_digest(),
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
            }

            # Get the item ID first
            query = f"FileLeafRef eq '{old_name}'"
            response = requests.get(
                f"{self.site_url}/_api/web/lists/GetByTitle('{library_name}')/Items?$filter={query}",
                headers=headers
            )

            if response.status_code != 200:
                raise Exception(f"Failed to find file. Status code: {response.status_code}")

            data = response.json()
            if not data['d']['results']:
                raise Exception("File not found")

            item_id = data['d']['results'][0]['ID']

            # Update the filename
            update_url = f"{self.site_url}/_api/web/lists/GetByTitle('{library_name}')/Items({item_id})"
            update_data = {'FileLeafRef': new_name}

            response = requests.post(
                update_url,
                headers=headers,
                json={'__metadata': {'type': 'SP.Data.DocumentsItem'}, **update_data}
            )

            if response.status_code not in [200, 204]:
                raise Exception(f"Failed to rename file. Status code: {response.status_code}")

            logger.info(f"File renamed successfully from {old_name} to {new_name}")

        except Exception as e:
            logger.error(f"Failed to rename file {old_name}: {str(e)}")
            raise

    def _get_form_digest(self):
        """Get form digest value for POST requests"""
        try:
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json;odata=verbose'
            }

            response = requests.post(
                f"{self.site_url}/_api/contextinfo",
                headers=headers
            )

            if response.status_code == 200:
                return response.json()['d']['GetContextWebInformation']['FormDigestValue']
            else:
                raise Exception("Failed to get form digest")

        except Exception as e:
            logger.error(f"Failed to get form digest: {str(e)}")
            raise