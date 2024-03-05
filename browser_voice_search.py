import time

import pyautogui
import pyperclip
import speech_recognition as sr
import win32con
import win32gui


def get_chrome_windows():
    """
    Retrieves a list of Chrome windows.
    Returns:
        List of tuples containing window handles and titles.
    """

    def window_enum_handler(hwnd, result_list):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
            result_list.append((hwnd, win32gui.GetWindowText(hwnd)))

    chrome_windows = []
    win32gui.EnumWindows(window_enum_handler, chrome_windows)
    chrome_windows = [
        (hwnd, title) for hwnd, title in chrome_windows if "chrome" in title.lower()
    ]
    return chrome_windows


def switch_to_chrome_tab():
    """
    Switches to the first Chrome window, maximizes it, and restores it.
    """
    chrome_windows = get_chrome_windows()
    if chrome_windows:
        hwnd, _ = chrome_windows[0]  # Assuming the first Chrome window
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # Maximize the window
        win32gui.SetForegroundWindow(hwnd)


def listen_and_convert():
    """
    Listens for a voice command using the microphone.
    Returns:
        The recognized voice command.
    """
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for command...")
        recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio)
        print("You said:", command)
        return command
    except sr.UnknownValueError:
        print("Could not understand audio.")
    except sr.RequestError as e:
        print("Could not request results; {0}".format(e))


def main():
    """
    Main function to listen for voice commands and perform actions.
    """
    while True:
        command = listen_and_convert()
        if command:
            if "SHUTDOWN" in command.upper():
                print("Shutting down...")
                break  # Exit the loop and terminate the program
            else:
                pyperclip.copy(command)  # Copy the command to clipboard
                switch_to_chrome_tab()
                time.sleep(1)  # Wait for the window switch
                pyautogui.hotkey("ctrl", "t")  # Open a new tab
                time.sleep(1)  # Wait for the tab to open
                pyautogui.hotkey("ctrl", "v")  # Paste from clipboard
                pyautogui.press("enter")  # Press enter to perform the search
                time.sleep(2)  # Wait for the search to complete


if __name__ == "__main__":
    main()
