import os
import shutil
import subprocess
import time
import tkinter as tk
from tkinter import messagebox

def prompt_user():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo("Focus Required", "Please make sure Notepad is the active window.")
    root.destroy()

def close_notepad_and_save(filename):
    """
    Close Notepad if the specified file is open and save the file.
    """
    notepad_command = f'tasklist /v /fi "imagename eq notepad.exe" /fo csv | findstr "{filename}"'
    result = subprocess.run(notepad_command, shell=True, capture_output=True, text=True)

    if filename in result.stdout:
        print(f'{filename} is open in Notepad. Saving and closing it...')
        save_command = (
            'powershell -command "$wshell = New-Object -ComObject wscript.shell; '
            '$wshell.AppActivate(\'Notepad\'); Start-Sleep 2; '  # Increased sleep time
            '$wshell.SendKeys(\'^s\'); Start-Sleep 2; '  # Increased sleep time
            '$wshell.SendKeys(\'%{F4}\')"'
        )
        
        try:
            subprocess.run(save_command, shell=True)
            print(f'Sent keystrokes to save and close {filename}.')
        except Exception as e:
            print(f'Error sending keystrokes: {e}')

def move_file(src, dest):
    """
    Move the file from source to destination folder with retry if needed.
    """
    for attempt in range(3):  # Retry up to 3 times if the file is in use
        try:
            if os.path.exists(src):  # Check if the source file exists
                shutil.move(src, dest)
                print(f'Moved {src} to {dest} successfully.')
                return True  # Break if the move is successful
            else:
                print(f'File not found: {src}')
        except Exception as e:
            print(f'Attempt {attempt + 1}: Error occurred while moving file: {e}')
            time.sleep(2)  # Wait before retrying
    return False  # Return False if it fails after all attempts

def main():
    # Show prompt to user
    prompt_user()
    
    # Define source and destination paths
    src_file = r"src\new.txt"
    dest_folder = r"dest"

    # Ensure the destination folder exists
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    
    # Close Notepad and save the file if it's open
    close_notepad_and_save(os.path.basename(src_file))
    
    # Wait for a brief moment to ensure the file is fully saved and released
    time.sleep(5)  # Increased sleep time for stability
    
    # Move the file
    if move_file(src_file, os.path.join(dest_folder, os.path.basename(src_file))):
        print("File moved successfully. Exiting...")
    else:
        print("Failed to move file after multiple attempts.")

if __name__ == "__main__":
    main()  # Execute the main function
