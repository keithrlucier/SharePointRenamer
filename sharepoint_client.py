import msal
from office365.runtime.auth.token_response import TokenResponse
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import logging
import os
import requests
import webbrowser

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url):
        """Initialize SharePoint client without immediate authentication"""
        self.site_url = site_url
        self.ctx = None
        # Extract tenant name from site URL
        self.tenant = self._extract_tenant_from_url(site_url)

    def _extract_tenant_from_url(self, url):
        """Extract tenant name from SharePoint URL"""
        import re
        match = re.search(r'https://([^.]+)\.sharepoint\.com', url)
        if match:
            return match.group(1)
        raise ValueError("Invalid SharePoint URL format. Expected: https://<tenant>.sharepoint.com/...")

    def authenticate(self):
        """Initialize authentication using SharePoint delegated permissions"""
        try:
            logger.info("Starting SharePoint authentication process...")

            # Get Azure AD app credentials
            client_id = os.environ.get('AZURE_CLIENT_ID')
            client_secret = os.environ.get('AZURE_CLIENT_SECRET')

            if not client_id or not client_secret:
                raise ValueError("AZURE_CLIENT_ID and AZURE_CLIENT_SECRET must be set")

            # Define SharePoint-specific endpoints
            tenant_id = "f8cdef31-a31e-4b4a-93e4-5f571e91255a"
            authority_url = f"https://login.microsoftonline.com/{tenant_id}"
            resource = f"https://{self.tenant}.sharepoint.com"

            logger.info(f"Authenticating with SharePoint site: {self.site_url}")
            logger.info(f"Using authority URL: {authority_url}")

            # Create authentication context
            auth_ctx = AuthenticationContext(self.site_url)

            # Set up with client credentials
            if auth_ctx.acquire_token_for_user(client_id, client_secret):
                self.ctx = ClientContext(self.site_url, auth_ctx)

                try:
                    # Test connection
                    self.ctx.load(self.ctx.web)
                    self.ctx.execute_query()
                    logger.info("Successfully connected to SharePoint site")
                    return True
                except Exception as e:
                    logger.error(f"Failed to connect to SharePoint site: {str(e)}")
                    headers = {
                        'Accept': 'application/json;odata=verbose'
                    }
                    # Add diagnostic information
                    logger.error("Attempting direct REST API test...")
                    response = requests.get(f"{self.site_url}/_api/web", headers=headers)
                    logger.error(f"REST API test response: {response.status_code}")
                    logger.error(f"REST API response content: {response.text}")
                    raise
            else:
                raise Exception("Failed to acquire token for user")

        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            if "AADSTS7000229" in str(e):
                logger.error("Service Principal not found. Please ensure the app is properly registered in Azure AD")
            raise

    def get_libraries(self):
        """Get all document libraries in the site"""
        try:
            if not self.ctx:
                self.authenticate()

            libraries = self.ctx.web.lists.filter("BaseTemplate eq 101").get().execute_query()
            return [lib.title for lib in libraries]
        except Exception as e:
            logger.error(f"Failed to get libraries: {str(e)}")
            raise

    def get_files(self, library_name):
        """Get all files in a library"""
        try:
            if not self.ctx:
                self.authenticate()

            target_list = self.ctx.web.lists.get_by_title(library_name)
            items = target_list.items.get().execute_query()

            files = []
            for item in items:
                files.append({
                    'Id': item.properties.get('ID'),
                    'Name': item.properties.get('FileLeafRef'),
                    'Path': item.properties.get('FileRef')
                })
            return files
        except Exception as e:
            logger.error(f"Failed to get files from library {library_name}: {str(e)}")
            raise

    def rename_file(self, library_name, old_name, new_name):
        """Rename a file in SharePoint"""
        try:
            if not self.ctx:
                self.authenticate()

            target_list = self.ctx.web.lists.get_by_title(library_name)
            items = target_list.items.filter(f"FileLeafRef eq '{old_name}'").get().execute_query()

            if len(items) > 0:
                file_item = items[0]
                file_item.set_property('FileLeafRef', new_name)
                file_item.update().execute_query()
                logger.info(f"File renamed successfully from {old_name} to {new_name}")
            else:
                raise Exception("File not found")
        except Exception as e:
            logger.error(f"Failed to rename file {old_name}: {str(e)}")
            raise