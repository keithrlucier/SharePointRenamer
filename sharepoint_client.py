import msal
from office365.runtime.auth.token_response import TokenResponse
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import logging
import os

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
        """Authenticate using MSAL device code flow"""
        try:
            logger.info("Starting SharePoint authentication process...")

            # Get Azure AD app registration client ID
            client_id = os.environ.get('AZURE_CLIENT_ID')
            if not client_id:
                raise ValueError("AZURE_CLIENT_ID environment variable is not set")

            # Use common endpoint for multi-tenant apps
            authority = "https://login.microsoftonline.com/common"

            logger.info("Initializing MSAL public client application...")

            # Initialize MSAL app as public client
            app = msal.PublicClientApplication(
                client_id,
                authority=authority
            )

            # Define SharePoint-specific scopes with resource URL
            scopes = [f"https://{self.tenant}.sharepoint.com/.default"]

            logger.info(f"Initiating device code flow with scopes: {scopes}")

            # Start device code flow
            flow = app.initiate_device_flow(scopes)

            if "user_code" not in flow:
                logger.error("Could not create device flow")
                raise Exception(
                    f"Could not create device flow. Error: {flow.get('error_description', 'No error description')}"
                )

            logger.info("Device code flow initiated successfully")
            return flow

        except Exception as e:
            logger.error(f"Failed to initialize authentication: {str(e)}")
            raise

    def complete_authentication(self, flow):
        """Complete the device code authentication flow"""
        try:
            logger.info("Completing device code authentication...")

            # Get Azure AD app registration client ID
            client_id = os.environ.get('AZURE_CLIENT_ID')
            if not client_id:
                raise ValueError("AZURE_CLIENT_ID environment variable is not set")

            # Use common endpoint for multi-tenant apps
            authority = "https://login.microsoftonline.com/common"

            app = msal.PublicClientApplication(
                client_id,
                authority=authority
            )

            # Get token using device code flow
            result = app.acquire_token_by_device_flow(flow)

            logger.info(f"Token acquisition result status: {'Success' if 'access_token' in result else 'Failed'}")

            if "access_token" in result:
                logger.info("Access token acquired successfully")
                # Create token response
                token = TokenResponse(result)

                # Initialize SharePoint context with token
                logger.info("Initializing SharePoint context...")
                self.ctx = ClientContext(self.site_url).with_access_token(token.access_token)

                logger.info("Loading SharePoint web context...")
                self.ctx.load(self.ctx.web)
                self.ctx.execute_query()

                logger.info("SharePoint client initialized successfully")
                return True
            else:
                error_msg = result.get("error_description", "No error description available")
                logger.error(f"Failed to acquire token: {error_msg}")
                raise Exception(f"Failed to acquire token: {error_msg}")

        except Exception as e:
            logger.error(f"Failed to complete authentication: {str(e)}")
            raise

    def get_libraries(self):
        """Get all document libraries in the site"""
        try:
            if not self.ctx:
                raise Exception("Client not authenticated")

            libraries = self.ctx.web.lists.filter("BaseTemplate eq 101").get().execute_query()
            return [lib.title for lib in libraries]
        except Exception as e:
            logger.error(f"Failed to get libraries: {str(e)}")
            raise

    def get_files(self, library_name):
        """Get all files in a library"""
        try:
            if not self.ctx:
                raise Exception("Client not authenticated")

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
                raise Exception("Client not authenticated")

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