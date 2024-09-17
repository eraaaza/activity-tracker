import pandas as pd
import time
from datetime import datetime
from pynput import keyboard, mouse
import threading
import platform
import subprocess
import psutil
import queue
import sys

# Queue for storing events
event_queue = queue.Queue()

def get_active_window():
    if platform.system() == 'Windows':
        import win32gui
        import win32process
        hwnd = win32gui.GetForegroundWindow()
        pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid[-1])
        return process.name()
    elif platform.system() == 'Darwin':
        from AppKit import NSWorkspace
        return NSWorkspace.sharedWorkspace().activeApplication()['NSApplicationName']
    elif platform.system() == 'Linux':
        try:
            window_id = subprocess.check_output(['xdotool', 'getactivewindow']).strip()
            pid = subprocess.check_output(['xdotool', 'getwindowpid', window_id]).strip()
            process = psutil.Process(int(pid))
            return process.name()
        except Exception as e:
            return 'Unable to access Linux process'
    else:
        print("Unsupported OS... \nExiting program.")
        sys.exit()


def on_press(key):
    try:
        event = {
            'timestamp': datetime.now(),
            'event_type': 'keystroke',
            'key': str(key),
            'application': get_active_window()
        }
        event_queue.put(event)
    except Exception as e:
        print(f"Error in on_press: {e}")

def on_click(x, y, button, pressed):
    if pressed:
        try:
            event = {
                'timestamp': datetime.now(),
                'event_type': 'mouse_click',
                'button': str(button),
                'application': get_active_window()
            }
            event_queue.put(event)
        except Exception as e:
            print(f"Error in on_click: {e}")

def process_events():
    event_list = []
    while True:
        try:
            # Get event from queue
            event = event_queue.get(timeout=1)
            event_list.append(event)
        except queue.Empty:
            # If queue is empty, save events to CSV
            if event_list:
                df = pd.DataFrame(event_list)
                # Append to CSV file
                df.to_csv('activity_log.csv', mode='a', header=not pd.io.common.file_exists('activity_log.csv'), index=False)
                event_list = []
        except Exception as e:
            print(f"Error in process_events: {e}")

def main():
    # Start keyboard listener
    keyboard_listener = keyboard.Listener(on_press=on_press)
    keyboard_listener.start()

    # Start mouse listener
    mouse_listener = mouse.Listener(on_click=on_click)
    mouse_listener.start()

    # Start event processing thread
    event_thread = threading.Thread(target=process_events)
    event_thread.daemon = True
    event_thread.start()

    print("Tracking started... Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping listeners and saving remaining events...")
        # Stop listeners
        keyboard_listener.stop()
        mouse_listener.stop()

        # Process remaining events
        remaining_events = []
        while not event_queue.empty():
            remaining_events.append(event_queue.get())

        if remaining_events:
            df = pd.DataFrame(remaining_events)
            df.to_csv('activity_log.csv', mode='a', header=not pd.io.common.file_exists('activity_log.csv'), index=False)

        print("Activity log saved to 'activity_log.csv'.")

if __name__ == '__main__':
    main()