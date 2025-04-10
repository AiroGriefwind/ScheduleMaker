from os import path, makedirs, listdir, remove
from sys import executable, argv
from json import loads
from shutil import copy2
from tempfile import mkdtemp, gettempdir
from logging import getLogger
from zipfile import ZipFile
from datetime import datetime
from packaging.version import parse as parse_version
import requests
import webbrowser  # Add this import for opening URLs

# PyInstaller command to create a standalone executable
#pyinstaller --onefile --windowed --add-data "*.json;." --add-data "zh_TW.qm;." ui.py

# GitHub repository information
GITHUB_OWNER = "AiroGriefwind"
GITHUB_REPO = "ScheduleMaker"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
CURRENT_VERSION = "1.0.1"  # Initial version - this should be updated with each release

class Updater:
    def __init__(self, logger=None):
        self.logger = logger or getLogger(__name__)
        self.release_url = None  # Store the URL for later use
    
    def check_for_updates(self):
        """Check if a newer version is available on GitHub"""
        try:
            self.logger.info("Checking for updates...")
            
            # Get the latest release information from GitHub
            response = requests.get(f"{GITHUB_API_URL}/releases/latest")
            response.raise_for_status()
            
            latest_release = response.json()
            latest_version = latest_release['tag_name'].lstrip('v')
            
            # Compare versions
            if parse_version(latest_version) > parse_version(CURRENT_VERSION):
                self.logger.info(f"New version available: {latest_version} (current: {CURRENT_VERSION})")
                
                # Extract the URL from release notes if it's a Google Drive link
                release_notes = latest_release['body']
                self.release_url = self._extract_url_from_text(release_notes)
                
                return {
                    'version': latest_version,
                    'download_url': self._get_asset_download_url(latest_release),
                    'release_notes': release_notes,
                    'release_url': self.release_url  # Add the extracted URL
                }
            else:
                self.logger.info("No updates available.")
                return None
        except Exception as e:
            self.logger.error(f"Error checking for updates: {str(e)}")
            return None
    
    def _extract_url_from_text(self, text):
        """Extract URL from text (assuming it contains a Google Drive link)"""
        import re
        # Pattern to match URLs, especially Google Drive links
        url_pattern = r'https?://(?:drive\.google\.com|docs\.google\.com)[^\s)"]+'
        match = re.search(url_pattern, text)
        if match:
            return match.group(0)
        return None
    
    def open_release_url(self):
        """Open the release URL in the default web browser"""
        if self.release_url:
            try:
                self.logger.info(f"Opening URL: {self.release_url}")
                webbrowser.open(self.release_url)
                return True
            except Exception as e:
                self.logger.error(f"Error opening URL: {str(e)}")
                return False
        return False
    
    # Rest of your methods remain the same
    def _get_asset_download_url(self, release):
        """Get the download URL for the appropriate asset"""
        # Look for a zip file asset that matches our platform
        for asset in release['assets']:
            # You can customize this logic based on your asset naming convention
            if asset['name'].endswith('.zip'):
                return asset['browser_download_url']
        return None
    
    def download_update(self, download_url):
        """Download the update package"""
        try:
            self.logger.info(f"Downloading update from {download_url}")
            
            # Create a temporary directory for the download
            temp_dir = mkdtemp()
            zip_path = path.join(temp_dir, "update.zip")
            
            # Download the file
            response = requests.get(download_url, stream=True)
            response.raise_for_status()
            
            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            self.logger.info(f"Update downloaded to {zip_path}")
            return zip_path
        except Exception as e:
            self.logger.error(f"Error downloading update: {str(e)}")
            return None
    
    def apply_update(self, update_zip):
        """Apply the update by extracting the zip file"""
        try:
            self.logger.info("Applying update...")
            
            # Create a backup of the current application
            backup_dir = self._create_backup()
            
            # Extract the update zip
            with ZipFile(update_zip, 'r') as zip_ref:
                # Get the current application directory
                app_dir = path.dirname(path.abspath(__file__))
                zip_ref.extractall(app_dir)
            
            # Clean up the temporary file
            remove(update_zip)
            
            self.logger.info("Update applied successfully!")
            return True
        except Exception as e:
            self.logger.error(f"Error applying update: {str(e)}")
            # Attempt to restore from backup
            self._restore_backup(backup_dir)
            return False
    
    def _create_backup(self):
        """Create a backup of the current application"""
        try:
            app_dir = path.dirname(path.abspath(__file__))
            backup_dir = path.join(gettempdir(), f"app_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            
            # Create backup directory
            makedirs(backup_dir)
            
            # Copy all Python files and JSON data files
            for filename in listdir(app_dir):
                if filename.endswith('.py') or filename.endswith('.json'):
                    src_path = path.join(app_dir, filename)
                    dst_path = path.join(backup_dir, filename)
                    copy2(src_path, dst_path)
            
            self.logger.info(f"Backup created at {backup_dir}")
            return backup_dir
        except Exception as e:
            self.logger.error(f"Error creating backup: {str(e)}")
            return None
    
    def _restore_backup(self, backup_dir):
        """Restore the application from backup"""
        if not backup_dir or not path.exists(backup_dir):
            self.logger.error("No backup available to restore")
            return False
        
        try:
            self.logger.info(f"Restoring from backup {backup_dir}")
            app_dir = path.dirname(path.abspath(__file__))
            
            # Copy all files from backup to app directory
            for filename in listdir(backup_dir):
                src_path = path.join(backup_dir, filename)
                dst_path = path.join(app_dir, filename)
                copy2(src_path, dst_path)
            
            self.logger.info("Restoration complete")
            return True
        except Exception as e:
            self.logger.error(f"Error restoring backup: {str(e)}")
            return False
