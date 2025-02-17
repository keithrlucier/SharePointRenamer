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

            # Request token for SharePoint with correct scope format
            scope = f"https://{self.tenant}.sharepoint.com/.default"
            logger.info(f"Requesting token for scope: {scope}")

            try:
                token_response = credential.get_token(scope)
                self.access_token = token_response.token
                logger.info("Token acquired successfully")
            except Exception as token_error:
                logger.error(f"Failed to acquire token: {str(token_error)}")
                raise

            # Test connection with detailed debugging
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
            }

            test_url = f"{self.site_url}/_api/web"
            logger.info(f"Testing SharePoint connection at: {test_url}")
            logger.debug(f"Request headers: {headers}")

            try:
                response = requests.get(test_url, headers=headers)
                logger.info(f"Response status code: {response.status_code}")
                logger.debug(f"Response headers: {dict(response.headers)}")

                if response.status_code == 200:
                    logger.info("Successfully authenticated with SharePoint")
                    return True
                else:
                    error_content = response.text
                    logger.error(f"SharePoint API Response Status: {response.status_code}")
                    logger.error(f"SharePoint API Response Headers: {dict(response.headers)}")
                    logger.error(f"SharePoint API Response Content: {error_content}")

                    if response.status_code == 401:
                        logger.error("Authentication failed - invalid credentials or insufficient permissions")
                        logger.error("Please verify:")
                        logger.error("1. The Azure AD application has the correct permissions")
                        logger.error("2. Admin consent is granted for ALL required permissions")
                        logger.error("3. The application is properly configured for client credentials flow")

                    raise Exception(f"Failed to connect to SharePoint. Status code: {response.status_code}")

            except requests.exceptions.RequestException as req_error:
                logger.error(f"Request failed: {str(req_error)}")
                raise

        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            logger.error("Please ensure:")
            logger.error("1. The Azure AD application is properly registered")
            logger.error("2. Admin consent is granted for SharePoint API permissions")
            logger.error("3. The Client ID and Secret are correct")
            logger.error("4. The SharePoint site URL is correct")
            logger.error("5. The Tenant ID matches your Azure AD tenant")
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
                logger.error(f"Failed to get libraries. Response: {response.text}")
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
                logger.error(f"Failed to get files. Response: {response.text}")
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
                logger.error(f"Failed to find file. Response: {response.text}")
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
                logger.error(f"Failed to rename file. Response: {response.text}")
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
                logger.error(f"Failed to get form digest. Response: {response.text}")
                raise Exception("Failed to get form digest")

        except Exception as e:
            logger.error(f"Failed to get form digest: {str(e)}")
            raise