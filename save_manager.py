# save_manager.py
import json
import os
import datetime
from typing import List, Dict, Any, Optional
import shutil

class SaveManager:
    def __init__(self, saves_directory: str = "saves"):
        """Initialize the SaveManager with a directory for saves."""
        self.saves_directory = saves_directory
        # Create saves directory if it doesn't exist
        if not os.path.exists(saves_directory):
            os.makedirs(saves_directory)
            
    def save_schedule(self, 
                     availability_data: Dict, 
                     schedule_data: Optional[Dict] = None, 
                     description: str = "", 
                     start_date: Optional[datetime.date] = None,
                     end_date: Optional[datetime.date] = None) -> str:
        """
        Save the current schedule and availability data.
        
        Args:
            availability_data: The availability data to save
            schedule_data: The generated schedule data (if any)
            description: User description of this save
            start_date: Schedule start date
            end_date: Schedule end date
            
        Returns:
            save_id: The ID of the created save
        """
        # Generate a timestamp-based ID for the save
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        save_id = f"save_{timestamp}"
        
        # Create save metadata
        save_metadata = {
            "id": save_id,
            "created_at": datetime.datetime.now().isoformat(),
            "description": description,
            "start_date": start_date.isoformat() if start_date else None,
            "end_date": end_date.isoformat() if end_date else None,
        }
        
        # Create save data object
        save_data = {
            "metadata": save_metadata,
            "availability_data": availability_data,
            "schedule_data": schedule_data
        }
        
        # Save to file
        save_path = os.path.join(self.saves_directory, f"{save_id}.json")
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(save_data, f, indent=2, ensure_ascii=False)
            
        return save_id
    
    def get_all_saves(self) -> List[Dict]:
        """
        Get a list of all available saves with their metadata.
        
        Returns:
            List of save metadata dictionaries
        """
        saves = []
        
        # List all save files
        for filename in os.listdir(self.saves_directory):
            if filename.endswith('.json') and filename.startswith('save_'):
                file_path = os.path.join(self.saves_directory, filename)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        save_data = json.load(f)
                        if "metadata" in save_data:
                            saves.append(save_data["metadata"])
                except Exception as e:
                    print(f"Error loading save {filename}: {e}")
        
        # Sort by creation date (newest first)
        saves.sort(key=lambda x: x.get("created_at", ""), reverse=True)
        return saves
    
    def load_save(self, save_id: str) -> Dict:
        """
        Load a specific save by ID.
        
        Args:
            save_id: The ID of the save to load
            
        Returns:
            The complete save data
            
        Raises:
            FileNotFoundError: If the save doesn't exist
        """
        save_path = os.path.join(self.saves_directory, f"{save_id}.json")
        
        if not os.path.exists(save_path):
            raise FileNotFoundError(f"Save {save_id} not found")
            
        with open(save_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def delete_save(self, save_id: str) -> bool:
        """
        Delete a save by ID.
        
        Args:
            save_id: The ID of the save to delete
            
        Returns:
            True if successful, False otherwise
        """
        save_path = os.path.join(self.saves_directory, f"{save_id}.json")
        
        if not os.path.exists(save_path):
            return False
            
        try:
            os.remove(save_path)
            return True
        except Exception:
            return False
            
    def backup_save(self, save_id: str, backup_dir: str = "backups") -> bool:
        """
        Create a backup of a save file.
        
        Args:
            save_id: The ID of the save to backup
            backup_dir: Directory to store backups
            
        Returns:
            True if successful, False otherwise
        """
        # Create backup directory if it doesn't exist
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            
        save_path = os.path.join(self.saves_directory, f"{save_id}.json")
        backup_path = os.path.join(backup_dir, f"{save_id}_backup.json")
        
        if not os.path.exists(save_path):
            return False
            
        try:
            shutil.copy2(save_path, backup_path)
            return True
        except Exception:
            return False
