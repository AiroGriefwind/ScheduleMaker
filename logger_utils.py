import logging
import json
import os
import shutil
import zipfile
from datetime import datetime
# Import for global exception handling
import sys
import traceback
def setup_global_exception_handler():
    """Set up a global exception hook to catch and log unhandled exceptions"""
    original_hook = sys.excepthook
    
    def exception_hook(exc_type, exc_value, exc_traceback):
        # Log the exception
        error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        logging.critical(f"Unhandled exception: {error_msg}")
        
        # Call the original exception hook
        original_hook(exc_type, exc_value, exc_traceback)
    
    sys.excepthook = exception_hook

def setup_logging():
    """Initialize the logging system with proper configuration"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # Use timestamp in filename for unique logs
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f"app_log_{timestamp}.log")
    
    # Configure logging
    logging.basicConfig(
        filename=log_file, 
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # Set up global exception handler
    setup_global_exception_handler()
    
    logging.info("Logging system initialized")
    return log_file


def log_error(error_message, exception=None):
    """Log an error with optional exception details"""
    if exception:
        logging.error(f"{error_message}: {str(exception)}", exc_info=True)
    else:
        logging.error(error_message)

def log_info(message):
    """Log an informational message"""
    logging.info(message)

def create_data_package():
    """Create a zip file containing all JSON data files for debugging"""
    try:
        # Create timestamp for unique filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        package_name = f"data_package_{timestamp}"
        package_dir = package_name
        
        # Create temporary directory
        if not os.path.exists(package_dir):
            os.makedirs(package_dir)
        
        # Copy all JSON files to the package directory
        json_files = [f for f in os.listdir() if f.endswith('.json')]
        for file in json_files:
            shutil.copy(file, package_dir)
        
        # Add the latest log file if it exists
        log_dir = "logs"
        if os.path.exists(log_dir):
            log_files = sorted([f for f in os.listdir(log_dir) if f.endswith('.log')], reverse=True)
            if log_files:
                latest_log = os.path.join(log_dir, log_files[0])
                shutil.copy(latest_log, package_dir)
        
        # Create zip file
        zip_path = f"{package_name}.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(package_dir):
                for file in files:
                    zipf.write(
                        os.path.join(root, file),
                        os.path.relpath(os.path.join(root, file), package_dir)
                    )
        
        # Clean up temporary directory
        shutil.rmtree(package_dir)
        
        logging.info(f"Data package created: {zip_path}")
        return zip_path
    except Exception as e:
        logging.error(f"Failed to create data package: {str(e)}", exc_info=True)
        return None
