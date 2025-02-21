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
        """Log file rename operations to a log file in the library"""
        try:
            if not self.access_token:
                self.authenticate()

            logger.info(f"Starting log creation for rename: {original_name} -> {new_name}")

            # Get drive ID for the library
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Accept': 'application/json'
            }

            # First get drive ID
            logger.info("Getting drive ID for log creation")
            response = requests.get(drives_url, headers=headers)
            if response.status_code != 200:
                raise Exception(f"Failed to get drives for log creation. Status code: {response.status_code}")

            drive_id = None
            for drive in response.json().get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    logger.info(f"Found drive ID: {drive_id}")
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found for log creation")

            # Prepare log entry
            log_filename = "FileRenameLog.txt"
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"""
Rename Operation: {timestamp}
Original Name: {original_name}
New Name: {new_name}
Reason: {reason}
----------------------------------------
"""
            # First try to find existing log file
            logger.info("Searching for existing log file")
            search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='{log_filename}')"
            search_response = requests.get(search_url, headers=headers)

            content_headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'text/plain; charset=utf-8'
            }

            # Using a retry mechanism for file operations
            max_retries = 3
            retry_count = 0
            backoff_time = 1  # Starting backoff time in seconds

            while retry_count < max_retries:
                try:
                    if search_response.status_code == 200 and search_response.json().get('value'):
                        # File exists, append to it
                        logger.info("Found existing log file, appending content")
                        file_id = search_response.json()['value'][0]['id']

                        # Get current content
                        get_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
                        get_response = requests.get(get_url, headers=headers)

                        if get_response.status_code == 200:
                            existing_content = get_response.text
                            # Always append new entry, remove duplicate check
                            full_content = existing_content + log_entry

                            # Update the file
                            update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
                            logger.info("Updating existing log file")
                            update_response = requests.put(
                                update_url,
                                headers=content_headers,
                                data=full_content.encode('utf-8')
                            )

                            if update_response.status_code in [200, 201]:
                                logger.info("Successfully updated log file")
                                return True
                            else:
                                raise Exception(f"Failed to update log file. Status: {update_response.status_code}")
                        else:
                            raise Exception(f"Failed to read existing log file. Status: {get_response.status_code}")
                    else:
                        # Create new file
                        logger.info("Creating new log file")
                        create_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{log_filename}:/content"
                        create_response = requests.put(
                            create_url,
                            headers=content_headers,
                            data=log_entry.encode('utf-8')
                        )

                        if create_response.status_code in [200, 201]:
                            logger.info("Successfully created new log file")
                            return True
                        else:
                            raise Exception(f"Failed to create log file. Status: {create_response.status_code}")

                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        raise Exception(f"Failed to create/update log after {max_retries} retries: {str(e)}")

                    logger.warning(f"Retry {retry_count} of {max_retries} for log file operation: {str(e)}")
                    time.sleep(backoff_time)  # Wait before retry with exponential backoff
                    backoff_time *= 2  # Double the backoff time for next retry

            raise Exception("Max retries exceeded for log file operation")

        except Exception as e:
            logger.error(f"Failed to create/update log entry: {str(e)}")
            logger.error("Error details:", exc_info=True)
            raise

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

    def bulk_rename_files(self, library_name, rename_operations):
        """Enhanced bulk rename with proper batch processing"""
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

            logger.info(f"Getting drive ID for library: {library_name}")
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

            logger.info(f"Found drive ID: {drive_id}")

            # Process rename operations
            results = []
            for operation in rename_operations:
                try:
                    file_id = operation['file_id']
                    new_name = operation['new_name']
                    old_name = operation['old_name']

                    logger.info(f"Attempting to rename: {old_name} -> {new_name}")

                    # Update file metadata
                    update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
                    update_data = {'name': new_name}

                    response = requests.patch(update_url, headers=headers, json=update_data)

                    if response.status_code in [200, 201]:
                        logger.info(f"Successfully renamed {old_name} to {new_name}")

                        # Add to results first
                        result = {
                            'old_name': old_name,
                            'new_name': new_name,
                            'success': True,
                            'file_id': file_id,
                            'error': None
                        }
                        results.append(result)

                        # Then create log entry
                        try:
                            self._create_rename_log(
                                library_name,
                                old_name,
                                new_name,
                                reason="Bulk rename operation"
                            )
                        except Exception as log_error:
                            logger.error(f"Failed to create log entry for {old_name}: {str(log_error)}")
                            # Continue with next rename even if logging fails
                    else:
                        error_message = response.text
                        logger.error(f"Failed to rename {old_name}: {error_message}")
                        results.append({
                            'old_name': old_name,
                            'new_name': new_name,
                            'success': False,
                            'file_id': file_id,
                            'error': error_message
                        })
                except Exception as e:
                    logger.error(f"Error processing rename for {operation['old_name']}: {str(e)}")
                    results.append({
                        'old_name': operation['old_name'],
                        'new_name': operation['new_name'],
                        'success': False,
                        'file_id': operation['file_id'],
                        'error': str(e)
                    })

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
                    path_parts = parent_path.split('/')
                    folder_name = path_parts[-1] if path_parts else 'Root'
                    folders[parent_path] = {
                        'id': file.get('Id'),
                        'name': folder_name,
                        'path': parent_path
                    }

            return list(folders.values())

        except Exception as e:
            logger.error(f"Failed to getfolders: {str(e)}")
            raise