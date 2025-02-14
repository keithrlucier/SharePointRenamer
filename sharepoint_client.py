import logging
import os
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
import requests

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url):
        """Initialize SharePoint client without immediate authentication"""
        self.site_url = site_url
        self.ctx = None
        self.tenant = self._extract_tenant_from_url(site_url)
        logger.info(f"Initialized SharePoint client for site: {site_url}")

    def _extract_tenant_from_url(self, url):
        """Extract tenant name from SharePoint URL"""
        import re
        match = re.search(r'https://([^.]+)\.sharepoint\.com', url)
        if match:
            return match.group(1)
        raise ValueError("Invalid SharePoint URL format. Expected: https://<tenant>.sharepoint.com/...")

    def _test_connection(self):
        """Test SharePoint connection with detailed diagnostics"""
        try:
            self.ctx.load(self.ctx.web)
            self.ctx.execute_query()
            logger.info("Successfully connected to SharePoint")
            return True
        except Exception as e:
            logger.error(f"Connection test failed: {str(e)}")
            self._run_diagnostics()
            raise

    def _run_diagnostics(self):
        """Run diagnostic tests on SharePoint connection"""
        try:
            logger.info("Running SharePoint connection diagnostics...")

            # Test basic site accessibility
            response = requests.get(f"{self.site_url}/_api/web", 
                                  headers={'Accept': 'application/json;odata=verbose'})

            logger.info(f"Site accessibility test - Status code: {response.status_code}")
            if response.status_code == 401:
                logger.error("Authentication failed - Please verify:")
                logger.error("1. Azure AD app registration is properly configured")
                logger.error("2. Admin consent is granted for SharePoint permissions")
                logger.error("3. Client ID and Secret are correct")
            elif response.status_code == 404:
                logger.error("SharePoint site not found - Please verify the site URL")

            logger.info("Diagnostics completed")
        except Exception as e:
            logger.error(f"Diagnostics failed: {str(e)}")

    def authenticate(self):
        """Initialize authentication using SharePoint client credentials"""
        try:
            logger.info("Starting SharePoint authentication process...")

            # Get Azure AD app credentials
            client_id = os.environ.get('AZURE_CLIENT_ID')
            client_secret = os.environ.get('AZURE_CLIENT_SECRET')

            if not client_id or not client_secret:
                raise ValueError("AZURE_CLIENT_ID and AZURE_CLIENT_SECRET must be set")

            logger.info(f"Authenticating with SharePoint site: {self.site_url}")

            # Create client credentials
            credentials = ClientCredential(client_id, client_secret)

            # Initialize SharePoint client context with credentials
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)

            # Test the connection and run diagnostics if needed
            return self._test_connection()

        except Exception as e:
            logger.error(f"Authentication failed: {str(e)}")
            if "AADSTS700016" in str(e):
                logger.error("Application not found in tenant. Please verify the application is registered correctly.")
            elif "AADSTS7000229" in str(e):
                logger.error("Service principal missing. Please ensure admin consent is granted in Azure AD.")
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