import os
import shutil
import subprocess
import time

def close_notepad_and_save(filename):
    """
    Close Notepad if the specified file is open and save the file.
    """
    # Check if Notepad is open with the file
    notepad_command = f'tasklist /v /fi "imagename eq notepad.exe" /fo csv | findstr "{filename}"'
    result = subprocess.run(notepad_command, shell=True, capture_output=True, text=True)
    
    if filename in result.stdout:
        print(f'{filename} is open in Notepad. Saving and closing it...')
        # Simulate Ctrl + S to save the file and Alt + F4 to close Notepad
        save_command = (
            'powershell -command "$wshell = New-Object -ComObject wscript.shell; '
            '$wshell.AppActivate(\'Notepad\'); Start-Sleep 1; '
            '$wshell.SendKeys(\'^s\'); Start-Sleep 1; '
            '$wshell.SendKeys(\'%{F4}\')"'
        )
        subprocess.run(save_command, shell=True)
        time.sleep(3)  # Wait for Notepad to save and close
        print(f'{filename} has been saved and Notepad has been closed.')

def close_explorer_if_open():
    """
    Close File Explorer if it is open.
    """
    explorer_command = 'tasklist /fi "imagename eq explorer.exe" /fo csv | findstr /i "explorer.exe"'
    result = subprocess.run(explorer_command, shell=True, capture_output=True, text=True)
    
    if 'explorer.exe' in result.stdout:
        print('File Explorer is open. Closing it...')
        subprocess.run('taskkill /f /im explorer.exe', shell=True)
        time.sleep(2)  # Wait for Explorer to close
        
        # Restart File Explorer after closing it
        subprocess.run('explorer.exe', shell=True)
        print('File Explorer restarted.')

def move_file(src, dest):
    """
    Move the file from source to destination folder with retry if needed.
    """
    for attempt in range(3):  # Retry up to 3 times if the file is in use
        try:
            shutil.move(src, dest)
            print(f"Moved {src} to {dest} successfully.")
            return True  # Break if the move is successful
        except Exception as e:
            print(f"Attempt {attempt + 1}: Error occurred while moving file: {e}")
            time.sleep(2)  # Wait for 2 seconds before retrying
    return False  # Return False if it fails after all attempts

def main():
    # Define source and destination paths
    src_file = r"src\new.txt"
    dest_folder = r"dest"
    
    # Ensure the destination folder exists
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    
    # Construct the destination file path by joining the folder and filename
    dest_file = os.path.join(dest_folder, os.path.basename(src_file))
    
    # Close Notepad and save the file if it's open
    close_notepad_and_save(os.path.basename(src_file))
    
    # Close File Explorer if it is open
    close_explorer_if_open()
    
    # Wait for a brief moment to ensure the file is fully released
    time.sleep(2)
    
    # Move the file
    file_moved = move_file(src_file, dest_file)
    
    # Check if the file was moved successfully
    if file_moved:
        print("File moved successfully. Exiting...")
    else:
        print("Failed to move file after multiple attempts.")
    
if __name__ == "__main__":
    main()
