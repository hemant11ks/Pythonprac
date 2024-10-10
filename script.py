import os
import shutil
import time
import psutil
import subprocess

def is_file_in_use(filepath):
    """
    Check if the file is in use by any process (e.g., Notepad).
    """
    for proc in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            open_files = proc.info['open_files']
            if open_files:
                for file in open_files:
                    if file.path == filepath:
                        print(f"{filepath} is being used by {proc.info['name']} (PID: {proc.info['pid']})")
                        return proc.info['pid']  # Return the PID of the process using the file
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def close_notepad_if_open(filepath):
    """
    Close Notepad if it is using the specified file.
    """
    pid = is_file_in_use(filepath)
    if pid:
        print(f"Closing Notepad with PID: {pid}")
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(2)  # Give the system time to close the process
        print(f"Closed Notepad with PID: {pid}")

def move_file(src, dest):
    """
    Move the file from source to destination folder.
    """
    try:
        shutil.move(src, dest)
        print(f"Moved {src} to {dest} successfully.")
    except Exception as e:
        print(f"Error occurred while moving file: {e}")

def main():
    # Define source and destination paths
    src_file = r"src\new.txt"
    dest_folder = r"dest"
    dest_file = os.path.join(dest_folder, os.path.basename(src_file))
    
    # Ensure the destination folder exists
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    
    # Close Notepad if the file is open
    close_notepad_if_open(src_file)
    
    # Move the file
    move_file(src_file, dest_file)

if __name__ == "__main__":
    main()
