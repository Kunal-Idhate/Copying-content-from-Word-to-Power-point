import time
import pyautogui

def copy_content():
    # Move the mouse to the starting position of the content in Word/Google Docs
    pyautogui.moveTo(x=100, y=100, duration=1)

    # Perform the keyboard shortcut to select all the content (Ctrl + A)
    pyautogui.hotkey('ctrl', 'a')

    # Perform the keyboard shortcut to copy the content (Ctrl + C)
    pyautogui.hotkey('ctrl', 'c')

def paste_content():
    # Move the mouse to the desired location in PowerPoint/Google Slides
    pyautogui.moveTo(x=200, y=200, duration=1)

    # Perform the keyboard shortcut to paste the content (Ctrl + V)
    pyautogui.hotkey('ctrl', 'v')

# Wait for the user to switch to the Word/Google Docs window
time.sleep(5)

# Copy the content
copy_content()

# Wait for the user to switch to the PowerPoint/Google Slides window
time.sleep(5)

# Paste the content
paste_content()
