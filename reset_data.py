"""
FilterTrail - Reset filter data to start fresh
"""

import os
import json

def reset_filter_data():
    """Reset filter data file to start fresh"""
    data_file = "filter_data.json"
    backup_file = "filter_data_backup.json"
    
    # Create empty filter history
    empty_data = []
    
    # Save to main file
    with open(data_file, 'w') as f:
        json.dump(empty_data, f, indent=2)
    print(f"Reset {data_file} to empty array")
    
    # Also reset backup if it exists
    if os.path.exists(backup_file):
        with open(backup_file, 'w') as f:
            json.dump(empty_data, f, indent=2)
        print(f"Reset {backup_file} to empty array")
        
    print("\nFilter history has been reset. You can start monitoring with a clean slate.")

if __name__ == "__main__":
    reset_filter_data()