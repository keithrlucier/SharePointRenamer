import logging
import os
import re
from azure.identity import ClientSecretCredential
import requests
from urllib.parse import urlparse
import datetime
import time

logger = logging.getLogger(__name__)

class SharePointClient:
    def __init__(self, site_url):
        """Initialize SharePoint client without immediate authentication"""
        self.site_url = site_url.rstrip('/')  # Remove trailing slash
        self.ctx = None
        self.tenant = self._extract_tenant_from_url(site_url)
        self.site_path = self._extract_site_path(site_url)
        self.access_token = None
        logger.info(f"Initialized SharePoint client for site: {site_url}")

    def _extract_tenant_from_url(self, url):
        """Extract tenant name from SharePoint URL"""
        match = re.search(r'https://([^.]+)\.sharepoint\.com', url)
        if match:
            return match.group(1)
        raise ValueError("Invalid SharePoint URL format. Expected: https://<tenant>.sharepoint.com/...")

    def _extract_site_path(self, url):
        """Extract the site path from SharePoint URL"""
        parsed = urlparse(url)
        path = parsed.path.rstrip('/')
        if not path:
            return ''
        return path

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

            # Construct the site ID using the full path
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"

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

            # Construct the site ID using the full path
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"

            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"
            logger.info(f"Fetching libraries from: {drives_url}")

            response = requests.get(drives_url, headers=headers)

            if response.status_code == 200:
                data = response.json()
                libraries = [drive['name'] for drive in data.get('value', [])]
                logger.info(f"Found {len(libraries)} libraries")
                return libraries
            else:
                logger.error(f"Failed to get libraries. Response: {response.text}")
                raise Exception(f"Failed to get libraries. Status code: {response.status_code}")

        except Exception as e:
            logger.error(f"Failed to get libraries: {str(e)}")
            raise

    def get_files(self, library_name):
        """Get all files in a library recursively using Microsoft Graph API"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }

            # First get the drive ID for the library
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
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

            def get_items_recursive(folder_id='root'):
                """Recursively get all items from a folder"""
                items = []

                # Get items from current folder
                items_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
                logger.info(f"Fetching items from folder: {items_url}")

                try:
                    response = requests.get(items_url, headers=headers)
                    if response.status_code == 200:
                        folder_items = response.json().get('value', [])

                        for item in folder_items:
                            if 'file' in item:
                                # This is a file
                                items.append({
                                    'Id': item['id'],
                                    'Name': item['name'],
                                    'Path': item['webUrl'],
                                    'ParentPath': item.get('parentReference', {}).get('path', ''),
                                    'Type': 'file'
                                })
                            elif 'folder' in item:
                                # This is a folder, recurse into it
                                logger.info(f"Found folder: {item['name']}, recursing...")
                                folder_items = get_items_recursive(item['id'])
                                items.extend(folder_items)
                    else:
                        logger.error(f"Failed to get items. Status code: {response.status_code}")
                        logger.error(f"Response: {response.text}")
                except Exception as e:
                    logger.error(f"Error fetching items: {str(e)}")

                return items

            # Start recursive file enumeration from root
            logger.info(f"Starting recursive file enumeration for library: {library_name}")
            all_items = get_items_recursive()
            logger.info(f"Found {len(all_items)} total items in {library_name}")

            return all_items

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
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
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

            # Log the rename operation AFTER successful rename
            self._create_rename_log(library_name, old_name, new_name)

        except Exception as e:
            logger.error(f"Failed to rename file {old_name}: {str(e)}")
            raise

    def _get_full_path(self, library_name, file_path):
        """Calculate the full path length including SharePoint URL"""
        # Base SharePoint URL + Library + File path
        if isinstance(file_path, dict):
            # Handle when file_path is a file object
            path_components = []
            if file_path.get('ParentPath'):
                path_components.append(file_path['ParentPath'])
            if file_path.get('Name'):
                path_components.append(file_path['Name'])
            file_path = '/'.join(path_components)
        full_path = f"{self.site_url}/{library_name}/{file_path}"
        return full_path.rstrip('/')

    def _is_path_too_long(self, full_path, max_length=240):
        """Check if path length exceeds safe limit"""
        return len(full_path) > max_length

    def _create_rename_log(self, library_name, original_name, new_name, reason="Path too long"):
        """Create a simple local log entry for file rename operations"""
        try:
            # Create logs directory if it doesn't exist
            log_dir = "logs"
            os.makedirs(log_dir, exist_ok=True)

            log_file_path = os.path.join(log_dir, "rename_operations.txt")

            # Create a local log entry with timestamp
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"""[{timestamp}] {library_name}: Renamed "{original_name}" to "{new_name}" ({reason})\n"""

            # Ensure atomic write with proper file handle cleanup
            try:
                with open(log_file_path, "a", encoding='utf-8') as f:
                    f.write(log_entry)
                    f.flush()  # Force flush to disk
                    os.fsync(f.fileno())  # Ensure it's written to disk

                logger.info(f"Successfully logged rename operation to {log_file_path}")
                logger.info(f"Log entry: {log_entry.strip()}")
                return True

            except IOError as io_err:
                logger.error(f"IOError while writing to log file: {str(io_err)}")
                return False

        except Exception as e:
            logger.error(f"Failed to create/write to log file: {str(e)}")
            return False

    def scan_for_long_paths(self, library_name):
        """Scan library for files with problematic path lengths"""
        try:
            if not self.access_token:
                self.authenticate()

            files = self.get_files(library_name)
            problematic_files = []

            for file in files:
                # Check both filename length and full path length
                filename_length = len(file['Name'])
                full_path = self._get_full_path(library_name, file)
                full_path_length = len(full_path)

                if filename_length > 128 or full_path_length > 256:
                    problematic_files.append({
                        'id': file['Id'],
                        'name': file['Name'],
                        'path': file['Path'],
                        'filename_length': filename_length,
                        'full_path_length': full_path_length,
                        'full_path': full_path
                    })

            return problematic_files

        except Exception as e:
            logger.error(f"Failed to scan for long paths: {str(e)}")
            raise

    def _upload_log_to_library(self, library_name, log_content):
        """Upload the rename log to the current SharePoint library"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'text/plain'
            }

            # Get drive ID for the library
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers=headers)
            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            drive_id = None
            for drive in response.json().get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found")

            # Upload log file
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"rename_operations_{timestamp}.txt"
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_name}:/content"

            response = requests.put(
                upload_url,
                headers=headers,
                data=log_content.encode('utf-8')
            )

            if response.status_code not in [200, 201]:
                logger.error(f"Failed to upload log file. Response: {response.text}")
                return False

            logger.info(f"Successfully uploaded log file {file_name} to SharePoint")
            return True

        except Exception as e:
            logger.error(f"Failed to upload log to SharePoint: {str(e)}")
            return False

    def bulk_rename_files(self, library_name, rename_operations):
        """Bulk rename with proper logging for each operation"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            }

            # Get drive ID for the library
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers=headers)
            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            drive_id = None
            for drive in response.json().get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found")

            # Initialize log content for SharePoint
            sharepoint_log_content = ""

            # Process rename operations
            results = []
            for operation in rename_operations:
                try:
                    file_id = operation['file_id']
                    new_name = operation['new_name']
                    old_name = operation['old_name']

                    # First log the attempt
                    logger.info(f"Attempting to rename: {old_name} -> {new_name}")

                    # Update file metadata
                    update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
                    update_data = {'name': new_name}

                    response = requests.patch(update_url, headers=headers, json=update_data)

                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if response.status_code in [200, 201]:
                        log_entry = f"[{timestamp}] {library_name}: Renamed \"{old_name}\" to \"{new_name}\" (Bulk rename operation)\n"
                        sharepoint_log_content += log_entry

                        # Log locally as well
                        log_success = self._create_rename_log(
                            library_name,
                            old_name,
                            new_name,
                            reason="Bulk rename operation"
                        )

                        if not log_success:
                            logger.error(f"Failed to log rename operation for {old_name}")

                        results.append({
                            'old_name': old_name,
                            'new_name': new_name,
                            'success': True,
                            'file_id': file_id,
                            'error': None,
                            'logged': log_success
                        })

                        logger.info(f"Successfully renamed {old_name} to {new_name}")
                    else:
                        error_message = response.text
                        log_entry = f"[{timestamp}] {library_name}: Failed to rename \"{old_name}\" to \"{new_name}\" (Error: {error_message})\n"
                        sharepoint_log_content += log_entry

                        # Log the failed rename attempt locally
                        self._create_rename_log(
                            library_name,
                            old_name,
                            new_name,
                            reason=f"Failed: {error_message}"
                        )

                        results.append({
                            'old_name': old_name,
                            'new_name': new_name,
                            'success': False,
                            'file_id': file_id,
                            'error': error_message,
                            'logged': True
                        })
                        logger.error(f"Failed to rename {old_name}: {error_message}")

                except Exception as e:
                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    log_entry = f"[{timestamp}] {library_name}: Error processing \"{operation.get('old_name', 'unknown')}\" (Error: {str(e)})\n"
                    sharepoint_log_content += log_entry

                    logger.error(f"Error processing rename for {operation.get('old_name', 'unknown')}: {str(e)}")

                    # Log the exception locally
                    self._create_rename_log(
                        library_name,
                        operation.get('old_name', 'unknown'),
                        operation.get('new_name', 'unknown'),
                        reason=f"Exception: {str(e)}"
                    )

                    results.append({
                        'old_name': operation.get('old_name', 'unknown'),
                        'new_name': operation.get('new_name', 'unknown'),
                        'success': False,
                        'file_id': operation.get('file_id', 'unknown'),
                        'error': str(e),
                        'logged': True
                    })

            # Upload log file to SharePoint
            if sharepoint_log_content:
                self._upload_log_to_library(library_name, sharepoint_log_content)

            return results

        except Exception as e:
            logger.error(f"Failed to perform bulk rename: {str(e)}")
            raise

    def create_test_library(self):
        """Create a test library with sample data including problematic file paths"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            }

            # First, check if Test Library exists
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            logger.info(f"Checking for existing Test Library at {drives_url}")
            response = requests.get(drives_url, headers=headers)

            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            test_library_exists = False
            test_library_id = None

            for drive in response.json().get('value', []):
                if drive['name'] == "Test Library":
                    test_library_exists = True
                    test_library_id = drive['id']
                    logger.info("Found existing Test Library")
                    break

            if not test_library_exists:
                # Create Test Library using SharePoint site endpoint
                lists_url = f"https://graph.microsoft.com/v1.0/{site_id}/lists"
                library_data = {
                    "displayName": "Test Library",
                    "list": {
                        "template": "documentLibrary"
                    }
                }

                logger.info("Creating new Test Library...")
                response = requests.post(lists_url, headers=headers, json=library_data)

                if response.status_code != 200 and response.status_code != 201:
                    logger.error(f"Failed to create library. Response: {response.text}")
                    raise Exception(f"Failed to create Test Library. Status code: {response.status_code}")

                # Get the drive ID for the newly created library
                response = requests.get(drives_url, headers=headers)
                if response.status_code != 200:
                    raise Exception("Failed to get drives after creation")

                for drive in response.json().get('value', []):
                    if drive['name'] == "Test Library":
                        test_library_id = drive['id']
                        break

                if not test_library_id:
                    raise Exception("Could not find Test Library after creation")

                logger.info("Created Test Library successfully")

            # Create sample files with different naming patterns
            sample_files = [
                # Deep nested path with long filename
                {
                    "path": "Very Long Path Example/Subfolder Level 1/Subfolder Level 2/Subfolder Level 3/Deep Nested Files/MEMORANDUM OF LAW IN SUPPORT OF DEFENDANTS MOTION TO DISMISS PLAINTIFFS FIRST AMENDED COMPLAINT FOR LACK OF PERSONAL JURISDICTION AND IMPROPER VENUE OR IN THE ALTERNATIVE MOTION TO TRANSFER VENUE.docx",
                    "content": "This file has both a long path and a long filename."
                },
                # Very long filename with repetitive text
                {
                    "path": "IN THE MATTER OF CERTAIN WIRELESS DEVICES WITH 3G AND-OR 4G CAPABILITIES AND COMPONENTS THEREOF ORDER NO. 85- GRANTING COMPLAINANT INTERDIGITALS MOTION TO STRIKE PORTIONS OF THE EXPERT REPORT OF DR. JAMES OLIVIER BASED ON NEW CONTENTIONS.txt",
                    "content": "This is an example of a very long filename that might cause issues."
                },
                # Long filename with special characters
                {
                    "path": "NOTICE OF FILING DEFENDANTS RESPONSE TO PLAINTIFFS FIRST SET OF INTERROGATORIES AND REQUEST FOR PRODUCTION OF DOCUMENTS - EXHIBIT A - CONFIDENTIAL - UNDER SEAL.pdf",
                    "content": "Sample long filename document with special characters"
                },
                # Multiple similar files with long names
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Pleadings/FIRST AMENDED COMPLAINT FOR DAMAGES AND DEMAND FOR JURY TRIAL - VERSION 1 - FINAL - APPROVED BY CLIENT.pdf",
                    "content": "First version of the document"
                },
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Pleadings/FIRST AMENDED COMPLAINT FOR DAMAGES AND DEMAND FOR JURY TRIAL - VERSION 2 - FINAL - APPROVED BY PARTNER.pdf",
                    "content": "Second version of the document"
                },
                # Files with same prefix but different content
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Discovery/PLAINTIFFS FIRST SET OF INTERROGATORIES TO DEFENDANT - PART 1 OF 3 - QUESTIONS 1-50.pdf",
                    "content": "Part 1"
                },
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Discovery/PLAINTIFFS FIRST SET OF INTERROGATORIES TO DEFENDANT - PART 2 OF 3 - QUESTIONS 51-100.pdf",
                    "content": "Part 2"
                }
            ]

            root_url = f"https://graph.microsoft.com/v1.0/drives/{test_library_id}/root"

            # Create necessary folders first
            folders = {
                "CASE TYH911 DST STATE OF FLORIDA E2E",
                "CASE TYH911 DST STATE OF FLORIDA E2E/Pleadings",
                "CASE TYH911 DST STATE OF FLORIDA E2E/Discovery",
                "Very Long Path Example/Subfolder Level 1/Subfolder Level 2/Subfolder Level 3/Deep Nested Files"
            }

            for folder_path in folders:
                folder_url = f"{root_url}:/{folder_path}:"
                response = requests.patch(folder_url, headers=headers, json={"folder": {}})

                if response.status_code not in [200, 201]:
                    logger.warning(f"Failed to create folder {folder_path}")
                else:
                    logger.info(f"Created folder: {folder_path}")

            # Create the sample files
            for file_info in sample_files:
                file_url = f"{root_url}:/{file_info['path']}:/content"
                file_response = requests.put(
                    file_url,
                    headers={'Authorization': f'Bearer {self.access_token}', 'Content-Type': 'text/plain'},
                    data=file_info['content'].encode('utf-8')
                )

                if file_response.status_code not in [200, 201]:
                    logger.warning(f"Failed to create file {file_info['path']}")
                else:
                    logger.info(f"Created file: {file_info['path']}")

            return True

        except Exception as e:
            logger.error(f"Failed to create test library: {str(e)}")
            raise

    def move_file(self, library_name, file_id, target_folder_id):
        """Move a file to a different folder using Microsoft Graph API"""
        try:
            if not self.access_token:
                self.authenticate()

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            }

            # First get the drive ID for the library
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers=headers)
            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            drive_id = None
            for drive in response.json().get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found")

            # Move the file to target folder
            move_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
            move_data = {
                'parentReference': {
                    'id': target_folder_id
                }
            }

            response = requests.patch(move_url, headers=headers, json=move_data)

            if response.status_code not in [200, 201]:
                logger.error(f"Failed to move file. Response: {response.text}")
                raise Exception(f"Failed to move file. Status code: {response.status_code}")

            return response.json()

        except Exception as e:
            logger.error(f"Failed to move file: {str(e)}")
            raise

    def get_folders(self, library_name):
        """Get all folders in a library"""
        try:
            if not self.access_token:
                self.authenticate()

            files = self.get_files(library_name)
            folders = {}

            for file in files:
                parent_path = file.get('ParentPath', '')
                if parent_path and parent_path not in folders:
                    path_parts = parentpath.split('/')
                    folder_name = path_parts[-1] if path_parts else 'Root'
                    folders[parent_path] = {
                        'id': file.get('Id'),
                        'name': folder_name,
                        'path': parent_path
                    }

            return list(folders.values())

        except Exception as e:
            logger.error(f"Failed to get folders: {str(e)}")
            raise