import os
import shutil
import subprocess
import time

def close_notepad_and_save(filename):
    # Check if Notepad is open with the specified file
    notepad_command = f'tasklist /v /fi "imagename eq notepad.exe" /fo csv | findstr "{filename}"'
    result = subprocess.run(notepad_command, shell=True, capture_output=True, text=True)
    
    if filename in result.stdout:
        print(f'Saving and closing {filename} in Notepad...')
        # Save the file and close Notepad
        subprocess.run(
            'powershell -command "$wshell = New-Object -ComObject wscript.shell; '
            '$wshell.AppActivate(\'Notepad\'); Start-Sleep 1; '
            '$wshell.SendKeys(\'^s\'); Start-Sleep 1; '
            '$wshell.SendKeys(\'%{F4}\')"',
            shell=True
        )
        time.sleep(3)  # Wait for Notepad to close

def move_file(src, dest):
    # Move the file with retry
    for attempt in range(3):
        try:
            shutil.move(src, dest)
            print(f'Moved {src} to {dest}.')
            return True
        except Exception as e:
            print(f'Attempt {attempt + 1}: {e}')
            time.sleep(2)  # Wait before retrying
    return False

def main():
    src_file = r"src\new.txt"
    dest_folder = r"dest"
    
    os.makedirs(dest_folder, exist_ok=True)  # Create destination folder if it doesn't exist
    
    close_notepad_and_save(os.path.basename(src_file))
    
    time.sleep(2)  # Wait for file to be released
    if move_file(src_file, os.path.join(dest_folder, os.path.basename(src_file))):
        print("File moved successfully. Exiting...")

if __name__ == "__main__":
    main()
