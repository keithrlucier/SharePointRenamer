import logging
import os
import re
from azure.identity import ClientSecretCredential
import requests
from urllib.parse import urlparse
import datetime

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
            # Get drive ID for the library
            host_part = f"{self.tenant}.sharepoint.com"
            site_path = self.site_path if self.site_path else ''
            site_id = f"sites/{host_part}{site_path}"
            drives_url = f"https://graph.microsoft.com/v1.0/{site_id}/drives"

            response = requests.get(drives_url, headers={'Authorization': f'Bearer {self.access_token}'})
            if response.status_code != 200:
                raise Exception(f"Failed to get drives. Status code: {response.status_code}")

            drive_id = None
            for drive in response.json().get('value', []):
                if drive['name'] == library_name:
                    drive_id = drive['id']
                    break

            if not drive_id:
                raise Exception(f"Library '{library_name}' not found")

            # Create or append to rename log file
            log_filename = "FileRenameLog.txt"
            log_content = f"""
Rename Operation: {datetime.datetime.now()}
Original Name: {original_name}
New Name: {new_name}
Reason: {reason}
----------------------------------------
"""
            # Check if log file exists
            search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='{log_filename}')"
            response = requests.get(search_url, headers={'Authorization': f'Bearer {self.access_token}'})

            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'text/plain'
            }

            if response.status_code == 200 and response.json().get('value'):
                # File exists, append to it
                file_id = response.json()['value'][0]['id']
                update_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"

                # Get existing content
                get_content_response = requests.get(update_url, headers={'Authorization': f'Bearer {self.access_token}'})
                existing_content = get_content_response.text if get_content_response.status_code == 200 else ""

                # Append new content
                full_content = existing_content + log_content
                requests.put(update_url, headers=headers, data=full_content.encode('utf-8'))
            else:
                # Create new file
                create_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/FileRenameLog.txt:/content"
                requests.put(create_url, headers=headers, data=log_content.encode('utf-8'))

        except Exception as e:
            logger.error(f"Failed to log rename operation: {str(e)}")

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

            # Process rename operations in batches
            batch_size = 20
            results = []

            # Process operations in chunks
            for i in range(0, len(rename_operations), batch_size):
                batch_chunk = rename_operations[i:i + batch_size]
                logger.info(f"Processing batch {i//batch_size + 1} with {len(batch_chunk)} operations")

                batch_requests = []
                request_map = {}

                for idx, operation in enumerate(batch_chunk):
                    request_id = str(idx + 1)
                    file_id = operation['file_id']
                    new_name = operation['new_name']

                    # Each request in the batch needs the full resource URL
                    batch_requests.append({
                        "id": request_id,
                        "method": "PATCH",
                        "url": f"/drives/{drive_id}/items/{file_id}",
                        "body": {"name": new_name},
                        "headers": {
                            "Content-Type": "application/json"
                        }
                    })
                    request_map[request_id] = operation

                batch_payload = {
                    "requests": batch_requests
                }

                logger.info(f"Sending batch request with {len(batch_requests)} operations")
                batch_url = "https://graph.microsoft.com/v1.0/$batch"
                batch_response = requests.post(batch_url, headers=headers, json=batch_payload)

                if batch_response.status_code != 200:
                    logger.error(f"Batch request failed. Status: {batch_response.status_code}")
                    logger.error(f"Response: {batch_response.text}")
                    raise Exception(f"Batch request failed. Status: {batch_response.status_code}")

                batch_results = batch_response.json().get('responses', [])
                logger.info(f"Received {len(batch_results)} responses from batch operation")

                for result in batch_results:
                    request_id = result['id']
                    original_op = request_map[request_id]
                    status_code = result['status']
                    success = 200 <= status_code < 300

                    if success:
                        logger.info(f"Successfully renamed {original_op['old_name']} to {original_op['new_name']}")
                        try:
                            self._create_rename_log(library_name, original_op['old_name'], original_op['new_name'])
                        except Exception as log_error:
                            logger.warning(f"Failed to create rename log: {str(log_error)}")
                    else:
                        error_body = result.get('body', {})
                        error_message = error_body.get('error', {}).get('message', 'Unknown error')
                        logger.error(f"Failed to rename {original_op['old_name']}: {error_message}")

                    results.append({
                        'old_name': original_op['old_name'],
                        'new_name': original_op['new_name'],
                        'success': success,
                        'file_id': original_op['file_id'],
                        'error': None if success else error_message
                    })

            logger.info(f"Completed processing {len(results)} rename operations")
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

                if response.status_code == 403:
                    error_msg = """
                    Access Denied: Your Azure AD application needs additional permissions.

                    Please ensure your Azure AD application has these permissions:
                    1. Sites.ReadWrite.All
                    2. Sites.Manage.All

                    To add these permissions:
                    1. Go to Azure Portal -> App Registrations
                    2. Select your application
                    3. Click on 'API Permissions'
                    4. Add the permissions above
                    5. Click 'Grant admin consent'
                    """
                    logger.error(error_msg)
                    raise Exception(error_msg)
                elif response.status_code not in [200, 201]:
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

            # Create sample folder structure and files
            root_url = f"https://graph.microsoft.com/v1.0/drives/{test_library_id}/root"

            # Create sample folders with varying depths
            folders = [
                "CASE TYH911 DST STATE OF FLORIDA E2E",
                "CASE TYH911 DST STATE OF FLORIDA E2E/Pleadings and Court Documents",
                "CASE TYH911 DST STATE OF FLORIDA E2E/Discovery Requests",
                "CASE TYH911 DST STATE OF FLORIDA E2E/Expert Reports and Analysis",
                "Very Long Path Example/Subfolder Level 1/Subfolder Level 2/Subfolder Level 3/Deep Nested Files"
            ]

            for folder_path in folders:
                folder_url = f"{root_url}:/{folder_path}:"
                response = requests.patch(folder_url, headers=headers, json={"folder": {}})

                if response.status_code not in [200, 201]:
                    logger.warning(f"Failed to create folder {folder_path}")
                else:
                    logger.info(f"Created folder: {folder_path}")

            # Create sample files with different naming patterns
            sample_files = [
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Simple File.txt",
                    "content": "This is a simple test file."
                },
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Pleadings and Court Documents/NOTICE OF FILING DEFENDANTS RESPONSE TO PLAINTIFFS FIRST SET OF INTERROGATORIES AND REQUEST FOR PRODUCTION OF DOCUMENTS.pdf",
                    "content": "Sample long filename document"
                },
                {
                    "path": "CASE TYH911 DST STATE OF FLORIDA E2E/Expert Reports and Analysis/IN THE MATTER OF CERTAIN WIRELESS DEVICES WITH 3G AND-OR 4G CAPABILITIES AND COMPONENTS THEREOF ORDER NO. 85- GRANTING COMPLAINANT INTERDIGITALS MOTION TO STRIKE PORTIONS OF THE EXPERT REPORT OF DR. JAMES OLIVIER BASED ON NEW CONTENTIONS AND TO PRECL.txt",
                    "content": "This is an example of a very long filename that might cause issues."
                },
                {
                    "path": "Very Long Path Example/Subfolder Level 1/Subfolder Level 2/Subfolder Level 3/Deep Nested Files/MEMORANDUM OF LAW IN SUPPORT OF DEFENDANTS MOTION TO DISMISS PLAINTIFFS FIRST AMENDED COMPLAINT FOR LACK OF PERSONAL JURISDICTION AND IMPROPER VENUE OR IN THE ALTERNATIVE MOTION TO TRANSFER VENUE.docx",
                    "content": "This file has both a long path and a long filename."
                }
            ]

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