#!/usr/bin/env python3
"""
Kobo to SharePoint Direct Transfer Pipeline
Transfers media directly from Kobo to SharePoint without local storage.
Workflow: Fetch Kobo forms → Extract media URLs from CSV → Stream directly to SharePoint
"""

import os
import sys
import csv
import json
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from typing import List, Dict, Optional, Tuple
import time
from io import BytesIO
import re

# Load environment variables
load_dotenv()


class DirectMediaTransfer:
    """Handles direct transfer of media from Kobo to SharePoint"""
    
    def __init__(self):
        """Initialize with credentials for both platforms"""
        # Kobo credentials
        self.kobo_token = os.getenv("API_TOKEN")
        if not self.kobo_token:
            raise ValueError("API_TOKEN not found in .env file")
        
        # SharePoint credentials
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.site_id = os.getenv("SITE_ID")
        
        if not all([self.tenant_id, self.client_id, self.client_secret, self.site_id]):
            raise ValueError("Missing SharePoint credentials in .env file")
        
        self.site_id = self.site_id.strip("'\"")
        
        # Sessions for both platforms
        self.kobo_session = requests.Session()
        self.kobo_session.headers.update({
            'Authorization': f'Token {self.kobo_token}',
            'User-Agent': 'KoboSharePointDirectTransfer/1.0'
        })
        
        self.sp_session = requests.Session()
        self.sp_access_token = None
        self.sp_drive_id = None
        
        # Transfer statistics
        self.stats = {
            'total_media': 0,
            'transferred': 0,
            'failed': 0,
            'skipped': 0,
            'total_size': 0
        }
        
        # Track naming format for each form
        self.current_form_naming_format = "fallback"
    
    def authenticate_sharepoint(self):
        """Get SharePoint access token"""
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        token_data = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default'
        }
        
        response = requests.post(token_url, data=token_data)
        response.raise_for_status()
        self.sp_access_token = response.json()['access_token']
        self.sp_session.headers.update({'Authorization': f'Bearer {self.sp_access_token}'})
        
        # Get drive ID
        drives_url = f'https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives'
        drives_response = self.sp_session.get(drives_url)
        drives_response.raise_for_status()
        
        drives = drives_response.json().get('value', [])
        if drives:
            self.sp_drive_id = drives[0]['id']
            print(f"📚 Using SharePoint library: {drives[0].get('name')}")
        else:
            raise ValueError("No document libraries found in SharePoint")
    
    def create_sharepoint_folder(self, folder_path: str) -> bool:
        """Create folder structure in SharePoint"""
        try:
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{folder_path}"
            check_response = self.sp_session.get(url)
            
            if check_response.status_code == 200:
                return True
            
            # Create folder hierarchy
            parts = folder_path.split('/')
            current_path = ""
            
            for part in parts:
                if current_path:
                    current_path = f"{current_path}/{part}"
                else:
                    current_path = part
                
                # Try to create each level
                if current_path == folder_path:
                    parent_path = ("/".join(parts[:-1])
                                   if len(parts) > 1 else "")
                    folder_name = parts[-1]
                    
                    if parent_path:
                        create_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{parent_path}:/children"
                    else:
                        create_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root/children"
                    
                    folder_data = {
                        "name": folder_name,
                        "folder": {},
                        "@microsoft.graph.conflictBehavior": "replace"
                    }
                    
                    self.sp_session.post(create_url, json=folder_data)
            
            return True
            
        except Exception as e:
            print(f"⚠️ Error creating folder {folder_path}: {e}")
            return False
    
    def stream_media_to_sharepoint(self, media_url: str, sharepoint_path: str, 
                                  filename: str) -> bool:
        """Stream media directly from Kobo URL to SharePoint without local storage"""
        try:
            print(f"   🔄 Streaming: {filename}")
            
            # Download from Kobo into memory
            kobo_response = self.kobo_session.get(media_url, stream=True, timeout=30)
            kobo_response.raise_for_status()
            
            # Get content size if available
            content_length = kobo_response.headers.get('content-length')
            if content_length:
                size_mb = int(content_length) / (1024 * 1024)
                print(f"      Size: {size_mb:.2f} MB")
            
            # Stream directly to SharePoint
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{sharepoint_path}/{filename}:/content"
            
            # For small files (< 4MB), upload directly
            if content_length and int(content_length) < 4 * 1024 * 1024:
                # Read into memory and upload
                content = kobo_response.content
                upload_response = self.sp_session.put(upload_url, data=content)
                
                if upload_response.status_code in [200, 201]:
                    self.stats['transferred'] += 1
                    self.stats['total_size'] += len(content)
                    print(f"      ✅ Transferred successfully")
                    return True
                else:
                    print(f"      ❌ SharePoint upload failed: {upload_response.status_code}")
                    self.stats['failed'] += 1
                    return False
            
            else:
                # For larger files, use chunked upload
                # Create upload session
                session_data = {
                    "item": {
                        "@microsoft.graph.conflictBehavior": "replace"
                    }
                }
                
                session_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{sharepoint_path}/{filename}:/createUploadSession"
                session_response = self.sp_session.post(session_url, json=session_data)
                
                if session_response.status_code != 200:
                    # Fall back to direct upload for medium-sized files
                    content = kobo_response.content
                    upload_response = self.sp_session.put(upload_url, data=content)
                    
                    if upload_response.status_code in [200, 201]:
                        self.stats['transferred'] += 1
                        self.stats['total_size'] += len(content)
                        print(f"      ✅ Transferred successfully")
                        return True
                    else:
                        print(f"      ❌ Upload failed: {upload_response.status_code}")
                        self.stats['failed'] += 1
                        return False
                
                # Upload in chunks
                upload_url = session_response.json()['uploadUrl']
                chunk_size = 320 * 1024  # 320 KB chunks
                file_size = int(content_length) if content_length else 0
                offset = 0
                
                for chunk in kobo_response.iter_content(chunk_size=chunk_size):
                    if chunk:
                        chunk_length = len(chunk)
                        headers = {
                            'Content-Length': str(chunk_length),
                            'Content-Range': f'bytes {offset}-{offset + chunk_length - 1}/{file_size}'
                        }
                        
                        chunk_response = requests.put(upload_url, data=chunk, headers=headers)
                        if chunk_response.status_code not in [200, 201, 202]:
                            print(f"      ❌ Chunk upload failed")
                            self.stats['failed'] += 1
                            return False
                        
                        offset += chunk_length
                
                self.stats['transferred'] += 1
                self.stats['total_size'] += file_size
                print(f"      ✅ Streamed successfully")
                return True
                
        except Exception as e:
            print(f"      ❌ Transfer error: {e}")
            self.stats['failed'] += 1
            return False
    
    def parse_attachments(self, attachments_str: str) -> List[Dict]:
        """Parse attachment JSON from CSV with better error handling"""
        if not attachments_str or attachments_str.strip() == '' or attachments_str == 'nan':
            return []
        
        try:
            cleaned_str = str(attachments_str).strip()
            
            # Handle different formats of attachment data
            if cleaned_str.startswith('[') and cleaned_str.endswith(']'):
                # It's already a JSON array
                pass
            elif cleaned_str.startswith('"[') and cleaned_str.endswith(']"'):
                # It's a quoted JSON array
                cleaned_str = cleaned_str[1:-1]
            elif cleaned_str.startswith('"') and cleaned_str.endswith('"'):
                # It's a quoted string, remove outer quotes
                cleaned_str = cleaned_str[1:-1]
            
            # Replace escaped quotes
            cleaned_str = cleaned_str.replace('\\"', '"')
            cleaned_str = cleaned_str.replace('\\/', '/')
            
            # Try to parse as JSON
            try:
                # First try direct JSON parsing
                attachments = json.loads(cleaned_str)
            except json.JSONDecodeError:
                # If that fails, try replacing single quotes with double quotes
                cleaned_str = cleaned_str.replace("'", '"')
                try:
                    attachments = json.loads(cleaned_str)
                except json.JSONDecodeError:
                    # If still failing, try to extract URLs directly
                    return self.extract_urls_fallback(attachments_str)
            
            # Ensure it's a list
            if isinstance(attachments, dict):
                attachments = [attachments]
            elif not isinstance(attachments, list):
                return []
            
            return attachments
            
        except Exception as e:
            # Last resort: try to extract URLs directly from the string
            return self.extract_urls_fallback(attachments_str)
    
    def extract_urls_fallback(self, text: str) -> List[Dict]:
        """Fallback method to extract URLs when JSON parsing fails"""
        
        urls = []
        # Look for URLs in the text
        url_pattern = r'https?://[^\s,\'"}\]]+(?:\.[^\s,\'"}\]]+)+'
        found_urls = re.findall(url_pattern, str(text))
        
        for url in found_urls:
            # Clean up the URL
            url = url.rstrip('\\').rstrip('"').rstrip("'")
            
            # Create a basic attachment structure
            if any(ext in url.lower() for ext in ['.jpg', '.jpeg', '.png', '.gif', '.pdf', '.mp4']):
                urls.append({
                    'download_url': url,
                    'filename': url.split('/')[-1].split('?')[0] or 'media_file'
                })
        
        return urls
    
    def sanitize_filename(self, filename: str) -> str:
        """Sanitize filename for safe file system use"""
        # Remove or replace unsafe characters
        safe_filename = re.sub(r'[^\w.-]', '_', filename)
        # Limit length
        if len(safe_filename) > 100:
            name, ext = os.path.splitext(safe_filename)
            safe_filename = name[:95] + ext
        return safe_filename
    
    def get_download_url(self, attachment: Dict) -> Optional[str]:
        """Get the best download URL from attachment data"""
        url_keys = ['download_large_url', 'download_url', 'download_medium_url']
        for key in url_keys:
            if key in attachment and attachment[key]:
                return attachment[key]
        return None
    
    def create_filename(self, row_data: Dict, row_num: int, base_filename: str,
                       date_column: str = "Date", type_column: str = "Receipt_Type") -> str:
        """Generate custom filename based on date and receipt type like the kobo-sharepoint file"""
        try:
            # Extract date from the specified column
            date_value = row_data.get(date_column, '')
            receipt_type = row_data.get(type_column, '')
            
            # Format date to YYYY-MM-DD format
            formatted_date = self.format_date(date_value)
            
            # Clean receipt type (remove spaces, special chars)
            clean_receipt_type = self.clean_text_for_filename(receipt_type)
            
            # Get file extension from base filename
            file_ext = os.path.splitext(base_filename)[1] or '.jpg'
            
            # Create custom filename: date_receiptType_rowNum.ext
            if formatted_date and clean_receipt_type:
                custom_name = f"{formatted_date}_{clean_receipt_type}_{row_num}{file_ext}"
                # Store the naming format for this form
                self.current_form_naming_format = "custom"
            else:
                # Fallback to original naming if data is missing
                custom_name = f"row{row_num}_{base_filename}"
                self.current_form_naming_format = "fallback"
            
            # Ensure filename is safe for filesystem
            safe_name = "".join(c for c in custom_name if c.isalnum() or c in "._-")
            return safe_name
            
        except Exception as e:
            # Fallback to simple naming
            self.current_form_naming_format = "fallback"
            return f"row{row_num}_{base_filename}"
    
    def format_date(self, date_str: str) -> str:
        """Format date string to YYYY-MM-DD format"""
        if not date_str:
            return ""
        try:
            # Handle common date formats
            date_formats = [
                "%Y-%m-%d",      # 2025-06-26
                "%m/%d/%Y",      # 06/26/2025
                "%d/%m/%Y",      # 26/06/2025
                "%Y-%m-%d %H:%M:%S",  # 2025-06-26 10:30:00
                "%Y-%m-%dT%H:%M:%S",  # 2025-06-26T10:30:00
                "%Y-%m-%dT%H:%M:%S.%f",  # ISO format with microseconds
            ]
            
            for fmt in date_formats:
                try:
                    from datetime import datetime
                    dt = datetime.strptime(date_str.strip(), fmt)
                    return dt.strftime("%Y-%m-%d")
                except ValueError:
                    continue
            
            # If no format matches, try to extract YYYY-MM-DD pattern
            match = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_str)
            if match:
                return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
            
            return ""
            
        except Exception as e:
            return ""
    
    def clean_text_for_filename(self, text: str) -> str:
        """Clean text to be safe for filenames"""
        if not text:
            return ""
        # Replace spaces with underscores and remove special characters
        cleaned = text.replace(' ', '_').replace('/', '_').replace('\\', '_')
        # Keep only alphanumeric and safe characters
        cleaned = "".join(c for c in cleaned if c.isalnum() or c in "_-")
        return cleaned
    
    def fetch_form_data_safe(self, form_uid: str) -> List[Dict]:
        """Fetch form data using the export endpoint which is more reliable"""
        try:
            # First try the standard data endpoint
            data_url = f"https://kf.kobotoolbox.org/api/v2/assets/{form_uid}/data.json"
            
            response = self.kobo_session.get(data_url)
            
            if response.status_code == 200:
                try:
                    # Try parsing the response
                    data = response.json()
                    if isinstance(data, dict) and 'results' in data:
                        return data['results']
                    elif isinstance(data, list):
                        return data
                except json.JSONDecodeError:
                    pass
            
            # If that fails, try the export endpoint
            print("      Trying alternative data endpoint...")
            export_url = f"https://kf.kobotoolbox.org/api/v2/assets/{form_uid}/data/?format=json"
            
            response = self.kobo_session.get(export_url)
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    if isinstance(data, dict) and 'results' in data:
                        return data['results']
                    elif isinstance(data, list):
                        return data
                except json.JSONDecodeError:
                    pass
            
            # Last resort: try CSV export and convert
            print("      Trying CSV export as fallback...")
            csv_url = f"https://kf.kobotoolbox.org/api/v2/assets/{form_uid}/data.csv"
            response = self.kobo_session.get(csv_url)
            
            if response.status_code == 200:
                import io
                
                # Parse CSV
                csv_text = response.text
                csv_reader = csv.DictReader(io.StringIO(csv_text))
                data = list(csv_reader)
                return data
            
            return []
            
        except Exception as e:
            print(f"      ❌ Error fetching form data: {str(e)[:100]}")
            return []
    
    def process_form_media(self, form_uid: str, form_name: str, 
                          main_folder: str, existing_files: Dict[str, set] = None) -> Dict:
        """Process all media for a single form with incremental upload support"""
        print(f"\n📋 Processing form: {form_name}")
        
        # Reset naming format for this form
        self.current_form_naming_format = "fallback"
        
        try:
            # Fetch form data using safe method
            data = self.fetch_form_data_safe(form_uid)
            
            if not data:
                print("   No submissions found or could not fetch data")
                return {'processed': 0, 'failed': 0, 'skipped': 0}
            
            print(f"   Found {len(data)} submissions")
            
            # Check naming format by testing first submission
            if data:
                first_submission = data[0]
                test_filename = self.create_filename(first_submission, 1, "test.jpg")
                print(f"   📝 Naming format: {self.current_form_naming_format} (example: {test_filename})")
            
            # Create form folder in SharePoint
            safe_form_name = re.sub(r'[^\w\s-]', '_', form_name)
            form_folder = f"{main_folder}/{safe_form_name}"
            self.create_sharepoint_folder(form_folder)
            
            form_stats = {'processed': 0, 'failed': 0, 'skipped': 0}
            
            # Process each submission
            for row_num, submission in enumerate(data, 1):
                if not isinstance(submission, dict):
                    continue
                    
                print(f"\n   📝 Submission {row_num}/{len(data)}")
                
                # Look for media in all columns
                for col_name, col_value in submission.items():
                    try:
                        # Skip if no value
                        if not col_value:
                            continue
                        
                        value_str = str(col_value)
                        
                        # Skip if doesn't look like media
                        if not any(indicator in value_str.lower() for indicator in 
                                  ['http', 'attachment', '.jpg', '.jpeg', '.png', '.pdf']):
                            continue
                        
                        # Extract media URLs
                        media_urls = []
                        
                        # Try JSON parsing first
                        if '[' in value_str or '{' in value_str:
                            attachments = self.parse_attachments(value_str)
                            for att in attachments:
                                url = self.get_download_url(att)
                                if url:
                                    media_urls.append({
                                        'url': url,
                                        'filename': att.get('filename', f'media_{row_num}')
                                    })
                        
                        # Also try direct URL extraction
                        if not media_urls and 'http' in value_str:
                            url_pattern = r'https?://[^\s,\'"}\]]+(?:\.[^\s,\'"}\]]+)+'
                            found_urls = re.findall(url_pattern, value_str)
                            for url in found_urls:
                                if any(ext in url.lower() for ext in ['.jpg', '.jpeg', '.png', '.pdf', '.mp4']):
                                    media_urls.append({
                                        'url': url,
                                        'filename': url.split('/')[-1].split('?')[0] or f'media_{row_num}'
                                    })
                        
                        if media_urls:
                            print(f"      Found {len(media_urls)} media file(s) in: {col_name}")
                            
                            # Create column folder
                            col_folder_name = re.sub(r'[^\w-]', '_', col_name)[:50]
                            col_folder = f"{form_folder}/{col_folder_name}"
                            self.create_sharepoint_folder(col_folder)
                            
                            for i, media in enumerate(media_urls):
                                self.stats['total_media'] += 1
                                
                                # Generate filename using custom naming convention
                                base_name = self.sanitize_filename(media['filename'])
                                filename = self.create_filename(submission, row_num, base_name)
                                
                                # Check if file already exists (incremental upload)
                                if existing_files and self.is_file_already_uploaded(filename, safe_form_name, existing_files):
                                    print(f"      ⏭️ Skipping existing file: {filename}")
                                    form_stats['skipped'] += 1
                                    self.stats['skipped'] += 1
                                    continue
                                
                                # Stream to SharePoint
                                success = self.stream_media_to_sharepoint(
                                    media['url'], col_folder, filename
                                )
                                
                                if success:
                                    form_stats['processed'] += 1
                                else:
                                    form_stats['failed'] += 1
                                
                                time.sleep(0.5)
                                
                    except Exception as e:
                        # Continue with next column
                        continue
            
            return form_stats
            
        except Exception as e:
            print(f"   ❌ Error processing form: {str(e)[:200]}")
            return {'processed': 0, 'failed': 0}
    
    def get_existing_files_in_sharepoint(self, main_folder: str) -> Dict[str, set]:
        """Get list of existing files in SharePoint to avoid duplicates"""
        existing_files = {}
        
        try:
            print(f"🔍 Checking existing files in SharePoint folder: {main_folder}")
            
            # List contents of main folder
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{main_folder}:/children"
            response = self.sp_session.get(url)
            
            if response.status_code != 200:
                print(f"⚠️ Could not access folder {main_folder}, assuming it's new")
                return existing_files
            
            items = response.json().get('value', [])
            print(f"   Found {len(items)} items in main folder")
            
            for item in items:
                if item.get('folder'):
                    form_name = item.get('name', '')
                    print(f"   📁 Checking form folder: {form_name}")
                    form_files = set()
                    
                    # Get files in this form folder
                    form_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{main_folder}/{form_name}:/children"
                    form_response = self.sp_session.get(form_url)
                    
                    if form_response.status_code == 200:
                        form_items = form_response.json().get('value', [])
                        print(f"      Found {len(form_items)} items in {form_name}")
                        
                        for file_item in form_items:
                            item_name = file_item.get('name', '')
                            if file_item.get('folder'):
                                print(f"      📁 Found subfolder: {item_name}")
                                # Recursively check subfolders
                                subfolder_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{main_folder}/{form_name}/{item_name}:/children"
                                subfolder_response = self.sp_session.get(subfolder_url)
                                if subfolder_response.status_code == 200:
                                    subfolder_items = subfolder_response.json().get('value', [])
                                    for subfile_item in subfolder_items:
                                        if not subfile_item.get('folder'):
                                            filename = subfile_item.get('name', '')
                                            form_files.add(filename)
                                            print(f"         📄 Found file in subfolder: {filename}")
                            else:
                                filename = file_item.get('name', '')
                                form_files.add(filename)
                                print(f"      📄 Found file: {filename}")
                    else:
                        print(f"      ❌ Could not access form folder {form_name}: {form_response.status_code}")
                        print(f"      Response: {form_response.text[:200]}")
                    
                    existing_files[form_name] = form_files
            
            total_files = sum(len(files) for files in existing_files.values())
            print(f"📊 Found {total_files} existing files across {len(existing_files)} forms")
            
            # Debug: show what we found
            for form_name, files in existing_files.items():
                if files:
                    print(f"   📋 {form_name}: {len(files)} files")
                    for filename in list(files)[:3]:  # Show first 3 files
                        print(f"      - {filename}")
                    if len(files) > 3:
                        print(f"      ... and {len(files) - 3} more")
            
            return existing_files
            
        except Exception as e:
            print(f"⚠️ Error checking existing files: {e}")
            import traceback
            traceback.print_exc()
            return existing_files
    
    def is_file_already_uploaded(self, filename: str, form_name: str, existing_files: Dict[str, set]) -> bool:
        """Check if a file already exists in SharePoint"""
        if form_name in existing_files:
            # Direct match
            if filename in existing_files[form_name]:
                print(f"      🔍 Found exact duplicate: {filename} in {form_name}")
                return True
            
            # For custom naming format, also check by row number and extension
            if self.current_form_naming_format == "custom":
                # Extract row number and extension from filename
                # Format: date_receiptType_rowNum.ext
                parts = filename.split('_')
                if len(parts) >= 3:
                    try:
                        row_num_part = parts[-1]  # Last part should be rowNum.ext
                        row_num = row_num_part.split('.')[0]  # Remove extension
                        file_ext = '.' + row_num_part.split('.')[1] if '.' in row_num_part else ''
                        
                        # Check if any existing file has the same row number and extension
                        for existing_file in existing_files[form_name]:
                            if existing_file.endswith(f"_{row_num}{file_ext}"):
                                print(f"      🔍 Found duplicate by row number: {existing_file} matches {filename}")
                                return True
                    except:
                        pass
            
            # For fallback naming format, check by row number and also exact match
            elif self.current_form_naming_format == "fallback":
                # Extract row number from filename: row{num}_{filename}
                if filename.startswith('row'):
                    try:
                        row_part = filename.split('_')[0]  # row{num}
                        row_num = row_part[3:]  # Remove 'row' prefix
                        
                        # Check if any existing file has the same row number
                        for existing_file in existing_files[form_name]:
                            if existing_file.startswith(f"row{row_num}_"):
                                print(f"      🔍 Found duplicate by row number: {existing_file} matches {filename}")
                                return True
                        
                        # Also check for exact filename match (in case of multiple files with same name)
                        if filename in existing_files[form_name]:
                            print(f"      🔍 Found exact duplicate: {filename}")
                            return True
                            
                    except Exception as e:
                        print(f"      ⚠️ Error checking fallback duplicates: {e}")
                        pass
        
        return False
    
    def verify_transfer(self, main_folder: str):
        """Verify files were transferred to SharePoint"""
        print("\n🔍 Verifying transfer...")
        
        try:
            # List contents of main folder
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.sp_drive_id}/root:/{main_folder}:/children"
            response = self.sp_session.get(url)
            
            if response.status_code != 200:
                print("❌ Could not access upload folder")
                return
            
            items = response.json().get('value', [])
            folder_count = sum(1 for item in items if item.get('folder'))
            
            print(f"✅ Verification complete:")
            print(f"   📁 Form folders created: {folder_count}")
            print(f"   📊 Files transferred: {self.stats['transferred']}")
            print(f"   💾 Total size: {self.stats['total_size'] / (1024*1024):.2f} MB")
            
            if self.stats['failed'] > 0:
                print(f"   ⚠️ Failed transfers: {self.stats['failed']}")
            
        except Exception as e:
            print(f"❌ Verification error: {e}")


class KoboSharePointDirectPipeline:
    """Main pipeline for direct transfer without local storage"""
    
    def __init__(self):
        """Initialize the pipeline"""
        self.transfer = DirectMediaTransfer()
        
        # Kobo API settings
        self.kobo_api_url = "https://kf.kobotoolbox.org/api/v2/assets/?format=json"
    
    def fetch_forms(self) -> pd.DataFrame:
        """Fetch available Kobo forms"""
        print("🔍 Fetching Kobo forms...")
        
        response = self.transfer.kobo_session.get(self.kobo_api_url)
        if response.status_code == 200:
            data = response.json()
            assets = data.get("results", [])
            print(f"✅ Found {len(assets)} forms")
            return pd.DataFrame(assets)
        else:
            print(f"❌ Failed to fetch forms")
            return pd.DataFrame()
    
    def select_forms(self, forms_df: pd.DataFrame) -> List[Dict]:
        """Interactive form selection"""
        if forms_df.empty:
            return []
        
        print("\n📋 Available forms:")
        for i, row in forms_df.iterrows():
            print(f"{i+1}. {row['name']} - Created: {row['date_created'][:10]}")
        
        print("\nEnter form numbers to transfer (comma-separated) or 'all':")
        choice = input("Your choice: ").strip()
        
        if choice.lower() == 'all':
            return forms_df[['uid', 'name']].to_dict('records')
        
        try:
            indices = [int(x.strip()) - 1 for x in choice.split(',')]
            selected = forms_df.iloc[indices][['uid', 'name']].to_dict('records')
            return selected
        except Exception:
            print("❌ Invalid selection")
            return []
    
    def run(self):
        """Execute the direct transfer pipeline"""
        print("🌟 Kobo to SharePoint Direct Transfer Pipeline")
        print("Transfers media directly without local storage")
        print("-" * 60)
        
        try:
            # Step 1: Authenticate with SharePoint
            print("\n🔐 Authenticating with SharePoint...")
            self.transfer.authenticate_sharepoint()
            print("✅ SharePoint authentication successful")
            
            # Step 2: Fetch and select Kobo forms
            forms_df = self.fetch_forms()
            selected_forms = self.select_forms(forms_df)
            
            if not selected_forms:
                print("❌ No forms selected")
                return
            
            # Step 3: Check for existing folder or create new one
            print("\n📁 Checking for existing upload folder...")
            
            # Find the most recent existing folder
            existing_folder = None
            latest_folder_date = None
            
            # List root items to find existing folders
            root_url = f"https://graph.microsoft.com/v1.0/sites/{self.transfer.site_id}/drives/{self.transfer.sp_drive_id}/root/children"
            root_response = self.transfer.sp_session.get(root_url)
            
            if root_response.status_code == 200:
                root_items = root_response.json().get('value', [])
                print(f"   📁 Found {len(root_items)} items in root folder")
                
                for item in root_items:
                    if item.get('folder') and item.get('name', '').startswith('KoboMedia_Direct_'):
                        folder_name = item.get('name', '')
                        print(f"   📁 Found Kobo folder: {folder_name}")
                        
                        # Extract date from folder name: KoboMedia_Direct_YYYYMMDD_HHMMSS
                        try:
                            # Split by underscore and get the date part (YYYYMMDD)
                            parts = folder_name.split('_')
                            if len(parts) >= 3:
                                folder_date = parts[2]  # YYYYMMDD part
                                if len(folder_date) == 8 and folder_date.isdigit():
                                    # Convert to datetime for comparison
                                    folder_datetime = datetime.strptime(folder_date, '%Y%m%d')
                                    print(f"      📅 Parsed date: {folder_datetime.strftime('%Y-%m-%d')}")
                                    
                                    if latest_folder_date is None or folder_datetime > latest_folder_date:
                                        latest_folder_date = folder_datetime
                                        existing_folder = folder_name
                                        print(f"      ✅ Selected as latest folder")
                                else:
                                    print(f"      ⚠️ Invalid date format: {folder_date}")
                            else:
                                print(f"      ⚠️ Invalid folder name format")
                        except Exception as e:
                            print(f"   ⚠️ Could not parse folder date from {folder_name}: {e}")
                            continue
            
            if existing_folder:
                print(f"📁 Using existing folder: {existing_folder}")
                print(f"   📅 Folder date: {latest_folder_date.strftime('%Y-%m-%d')}")
                main_folder = existing_folder
            else:
                # Create new folder
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                main_folder = f"KoboMedia_Direct_{timestamp}"
                print(f"📁 Creating new SharePoint folder: {main_folder}")
                self.transfer.create_sharepoint_folder(main_folder)
            
            # Step 4: Get existing files for incremental upload
            existing_files = self.transfer.get_existing_files_in_sharepoint(main_folder)
            
            # Step 5: Process each form with incremental upload
            print(f"\n🚀 Starting incremental transfer of {len(selected_forms)} forms...")
            
            for form in selected_forms:
                form_stats = self.transfer.process_form_media(
                    form['uid'], 
                    form['name'],
                    main_folder,
                    existing_files
                )
                print(f"   ✅ {form['name']}: {form_stats['processed']} transferred, "
                      f"{form_stats['failed']} failed, {form_stats['skipped']} skipped")
            
            # Step 5: Verify and summarize
            self.transfer.verify_transfer(main_folder)
            
            print("\n" + "=" * 60)
            print("📊 TRANSFER COMPLETE")
            print("=" * 60)
            print(f"Total media found: {self.transfer.stats['total_media']}")
            print(f"Successfully transferred: {self.transfer.stats['transferred']}")
            print(f"Failed transfers: {self.transfer.stats['failed']}")
            print(f"Skipped (already exists): {self.transfer.stats['skipped']}")
            print(f"Total size transferred: {self.transfer.stats['total_size'] / (1024*1024):.2f} MB")
            print(f"SharePoint location: {main_folder}/")
            print("=" * 60)
            
            print("\n✅ Direct transfer completed successfully!")
            
        except KeyboardInterrupt:
            print("\n⚠️ Transfer interrupted by user")
        except Exception as e:
            print(f"\n❌ Transfer failed: {e}")
            import traceback
            traceback.print_exc()


def main():
    """Main entry point"""
    # Load and validate environment variables
    load_dotenv()
    
    required_vars = ['API_TOKEN', 'TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'SITE_ID']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print("❌ Missing required environment variables:")
        for var in missing_vars:
            print(f"   - {var}")
        print("\nPlease ensure all variables are set in your .env file.")
        sys.exit(1)
    
    # Run the pipeline
    pipeline = KoboSharePointDirectPipeline()
    pipeline.run()


if __name__ == "__main__":
    main()