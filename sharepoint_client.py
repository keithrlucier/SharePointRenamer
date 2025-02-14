from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import logging

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url, username, password):
        """Initialize SharePoint client"""
        try:
            self.ctx = ClientContext(site_url).with_credentials(
                UserCredential(username, password)
            )
            self.ctx.load(self.ctx.web)
            self.ctx.execute_query()
            logger.info("SharePoint client initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize SharePoint client: {str(e)}")
            raise

    def get_libraries(self):
        """Get all document libraries in the site"""
        try:
            libraries = self.ctx.web.lists.filter("BaseTemplate eq 101").get().execute_query()
            return [lib.title for lib in libraries]
        except Exception as e:
            logger.error(f"Failed to get libraries: {str(e)}")
            raise

    def get_files(self, library_name):
        """Get all files in a library"""
        try:
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
