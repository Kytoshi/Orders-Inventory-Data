import os
import time
import shutil
from logger import logger
from datetime import datetime, timedelta
from helpers import subtract_one_business_day, get_current_dir
import send2trash

def find_and_copy_file(source_folder, destination_folder, file_prefix):
    """
    Finds the latest file in the source_folder that starts with file_prefix, 
    renames it with the current date, and copies it to the destination_folder.
    If a file with the same name exists, increments with (1), (2), etc.
    """
    try:
        files = [f for f in os.listdir(source_folder) if f.startswith(file_prefix)]
        
        if not files:
            logger.info(f"No files found with prefix '{file_prefix}' in {source_folder}")
            return None
    except Exception as e:
        logger.error(f"Error accessing source folder: {e}")
        return None
    
    # Get the most recent file based on modification time
    files.sort(key=lambda x: os.path.getmtime(os.path.join(source_folder, x)), reverse=True)
    latest_file = files[0]
    
    # Generate the new filename with current date
    file_name, file_extension = os.path.splitext(latest_file)
    # base_new_name = f"{file_name}_{subtract_one_business_day(datetime.today().date())}"
    base_new_name = f"{file_name}_{datetime.today().date()}"
    
    source_path = os.path.join(source_folder, latest_file)
    
    # Find an available filename (increment if needed)
    destination_path = os.path.join(destination_folder, f"{base_new_name}{file_extension}")
    counter = 1
    
    while os.path.exists(destination_path):
        # Add (1), (2), etc. before the extension
        destination_path = os.path.join(
            destination_folder, 
            f"{base_new_name} ({counter}){file_extension}"
        )
        counter += 1
    
    try:
        # Ensure destination folder exists
        os.makedirs(destination_folder, exist_ok=True)
        
        shutil.copy2(source_path, destination_path)
        final_name = os.path.basename(destination_path)
        logger.info(f"✓ Copied '{latest_file}' to '{destination_folder}' as '{final_name}'")
        return destination_path
    except Exception as e:
        logger.error(f"✗ Error copying file: {e}")
        return None

def remove_old_files(folder_path):
    """
    Removes old data files and transfer them into the Recycle Bin/Trash.
    """
    target_prefixes = [
        "Billing Only", 
        "DailyReport Completed", 
        "DailyReport Incompletes",
        "MatShortageRpt",
    ]

    try:
        files = os.listdir(folder_path)
    except FileNotFoundError:
        logger.error(f"Folder does not exist: {folder_path}")
        return
    except Exception as e:
        logger.error(f"Error accessing folder: {folder_path}. Error: {e}")
        return

    for file in files:
        # Check if file matches any target prefix
        if any(file.startswith(prefix) for prefix in target_prefixes):
            file_path = os.path.join(folder_path, file)
            
            # Only delete files, not directories
            if os.path.isfile(file_path):
                try:
                    send2trash.send2trash(file_path)
                    logger.info(f"Moved to Trash: {file}")
                except FileNotFoundError:
                    logger.warning(f"File no longer exists (deleted by another process): {file}")
                except PermissionError:
                    logger.error(f"Permission denied deleting file: {file_path}")
                except Exception as e:
                    logger.error(f"Error moving file to Trash: {file_path}. Error: {e}")

def wait_for_download(file, timeout=300, after_time=None):
    """
    Wait for a file download to complete.

    Args:
        file: Filename prefix to search for
        timeout: Maximum time to wait in seconds
        after_time: If provided, only consider files modified after this timestamp (time.time())

    Returns:
        Path to the downloaded file
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        files = [f for f in os.listdir(get_current_dir()) if f.startswith(file)]
        if files:
            # Sort by modification time (newest first) to get the most recent download
            files_with_mtime = [
                (f, os.path.getmtime(os.path.join(get_current_dir(), f)))
                for f in files
            ]
            files_with_mtime.sort(key=lambda x: x[1], reverse=True)

            newest_file, mtime = files_with_mtime[0]
            file_path = os.path.join(get_current_dir(), newest_file)

            # Skip if file is older than required time (prevents picking up old files)
            if after_time and mtime < after_time:
                time.sleep(1)
                continue

            if file_path.endswith(".crdownload") or file_path.endswith(".part"):  # Temporary download files
                time.sleep(1)  # Wait and check again
            else:
                logger.info(f"Download complete: {file_path}")
                return file_path
        time.sleep(1)
    else:
        raise TimeoutError("File download timed out.")