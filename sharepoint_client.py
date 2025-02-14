import msal
from office365.runtime.auth.token_response import TokenResponse
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import logging

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url):
        """Initialize SharePoint client without immediate authentication"""
        self.site_url = site_url
        self.ctx = None

    def authenticate(self):
        """Authenticate using MSAL device code flow"""
        try:
            logger.info("Starting SharePoint authentication process...")

            # Azure AD app registration details
            client_id = "1b730954-1685-4b74-9bfd-dac224a7b894"  # Microsoft Graph PowerShell client ID
            authority = "https://login.microsoftonline.com/common"

            logger.info("Initializing MSAL application...")

            # Initialize MSAL app
            app = msal.PublicClientApplication(
                client_id,
                authority=authority
            )

            # Define required scopes - Using only Graph API scopes
            scopes = [
                "https://graph.microsoft.com/Sites.Read.All",
                "https://graph.microsoft.com/Sites.ReadWrite.All"
            ]

            logger.info("Initiating device code flow...")

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

            # Azure AD app registration details
            client_id = "1b730954-1685-4b74-9bfd-dac224a7b894"
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