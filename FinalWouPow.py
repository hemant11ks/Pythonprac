import os
import shutil
import subprocess
import time

def close_notepad_and_save(filename):
    # Check if Notepad is running with the specified file open
    notepad_command = f'tasklist /v /fi "imagename eq notepad.exe" /fo csv | findstr "{filename}"'
    result = subprocess.run(notepad_command, shell=True, capture_output=True, text=True)
    
    # If the specified file is found in Notepad, save and close it
    if filename in result.stdout:
        print(f'Saving and closing {filename} in Notepad...')
        # Use PowerShell to send Ctrl+S (save) and Alt+F4 (close) keystrokes to Notepad
        subprocess.run(
            'powershell -command "$wshell = New-Object -ComObject wscript.shell; '
            '$wshell.AppActivate(\'Notepad\'); Start-Sleep 1; '
            '$wshell.SendKeys(\'^s\'); Start-Sleep 1; '
            '$wshell.SendKeys(\'%{F4}\')"',
            shell=True
        )
        time.sleep(3)  # Wait for a few seconds to ensure Notepad closes

def move_file(src, dest):
    # Attempt to move the file from the source to the destination
    for attempt in range(3):  # Retry up to 3 times if there's an error
        try:
            shutil.move(src, dest)  # Move the file
            print(f'Moved {src} to {dest}.')  # Success message
            return True  # Exit the loop if successful
        except Exception as e:
            print(f'Attempt {attempt + 1}: {e}')  # Print error message
            time.sleep(2)  # Wait 2 seconds before retrying
    return False  # Return False if the move fails after all attempts

def main():
    # Define the source file path and destination folder
    src_file = r"src\new.txt"
    dest_folder = r"dest"
    
    # Create the destination folder if it doesn't already exist
    os.makedirs(dest_folder, exist_ok=True)  
    
    # Call the function to close Notepad and save the file if it is open
    close_notepad_and_save(os.path.basename(src_file))
    
    time.sleep(2)  # Wait briefly to ensure the file is released
    
    # Attempt to move the file and check if it was successful
    if move_file(src_file, os.path.join(dest_folder, os.path.basename(src_file))):
        print("File moved successfully. Exiting...")

if __name__ == "__main__":
    main()  # Execute the main function
