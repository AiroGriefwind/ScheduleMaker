import os
import sys
import json
import shutil
import tempfile
import logging
import zipfile
import requests
from datetime import datetime
from packaging import version

# GitHub repository information
GITHUB_OWNER = "your-github-username"
GITHUB_REPO = "your-repo-name"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
CURRENT_VERSION = "1.0.0"  # Initial version - this should be updated with each release

class Updater:
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
    
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
            if version.parse(latest_version) > version.parse(CURRENT_VERSION):
                self.logger.info(f"New version available: {latest_version} (current: {CURRENT_VERSION})")
                return {
                    'version': latest_version,
                    'download_url': self._get_asset_download_url(latest_release),
                    'release_notes': latest_release['body']
                }
            else:
                self.logger.info("No updates available.")
                return None
        except Exception as e:
            self.logger.error(f"Error checking for updates: {str(e)}")
            return None
    
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
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "update.zip")
            
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
            with zipfile.ZipFile(update_zip, 'r') as zip_ref:
                # Get the current application directory
                app_dir = os.path.dirname(os.path.abspath(__file__))
                zip_ref.extractall(app_dir)
            
            # Clean up the temporary file
            os.remove(update_zip)
            
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
            app_dir = os.path.dirname(os.path.abspath(__file__))
            backup_dir = os.path.join(tempfile.gettempdir(), f"app_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            
            # Create backup directory
            os.makedirs(backup_dir)
            
            # Copy all Python files and JSON data files
            for filename in os.listdir(app_dir):
                if filename.endswith('.py') or filename.endswith('.json'):
                    src_path = os.path.join(app_dir, filename)
                    dst_path = os.path.join(backup_dir, filename)
                    shutil.copy2(src_path, dst_path)
            
            self.logger.info(f"Backup created at {backup_dir}")
            return backup_dir
        except Exception as e:
            self.logger.error(f"Error creating backup: {str(e)}")
            return None
    
    def _restore_backup(self, backup_dir):
        """Restore the application from backup"""
        if not backup_dir or not os.path.exists(backup_dir):
            self.logger.error("No backup available to restore")
            return False
        
        try:
            self.logger.info(f"Restoring from backup {backup_dir}")
            app_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Copy all files from backup to app directory
            for filename in os.listdir(backup_dir):
                src_path = os.path.join(backup_dir, filename)
                dst_path = os.path.join(app_dir, filename)
                shutil.copy2(src_path, dst_path)
            
            self.logger.info("Restoration complete")
            return True
        except Exception as e:
            self.logger.error(f"Error restoring backup: {str(e)}")
            return False
