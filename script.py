import pygetwindow as gw
from pypresence import Presence
import time
import psutil
import win32process
from pystray import Icon as icon, Menu, MenuItem as item
import pystray
from PIL import Image
import threading

# Global flag to signal exit
exit_flag = False

# Function to run the system tray icon
def run_tray_icon():
    image = Image.open('keyboard.ico')
    menu = Menu(item('Quit', lambda: set_exit_flag()))
    tray_icon = pystray.Icon('notion-presence', image, 'notion-presence', menu)
    tray_icon.run()

# Function to set the exit/quit flag
def set_exit_flag():
    global exit_flag
    exit_flag = True

def get_window_process_id(hwnd):
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid

def find_notion_window():
    # Check for the process ids of "Notion.exe"
    processes_hit_list = []
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'Notion.exe':
            processes_hit_list.append(proc.info['pid'])

    # Find the window with the process id of "Notion.exe" then return the taskbar window of the application "Notion.exe"
    for window in gw.getAllWindows():
        if get_window_process_id(window._hWnd) in processes_hit_list:
            return window
    return None

# Find the Notion window by process name
notion_window = find_notion_window()

if notion_window:
    title = notion_window.title
else:
    title = "Notion is not running"

"""
You need to upload your image(s) here:
https://discordapp.com/developers/applications/<APP ID>/rich-presence/assets
"""

client_id = "861651444234846258"  # application ID from discord
RPC = Presence(client_id=client_id)
RPC.connect()
start_time = int(time.time())

# use the same name that you used when uploading the image
RPC.update(large_image="notion_1024", large_text="Notion",
           small_image="keeb_512", small_text="Taking notes", details=("Editing: " + title), state="In the app")

# ISSUE: only exits every 15 seconds interval after the exit flag is set
# Start the system tray icon in a separate thread (need this otherwise the system tray icon won't show up)
tray_thread = threading.Thread(target=run_tray_icon)
tray_thread.daemon = True  # Daemonize thread to exit when the main program exits
tray_thread.start()

# update the status every 15 seconds
while not exit_flag:
    notion_window = find_notion_window()

    if notion_window is None:
        if RPC is not None:
            RPC.close()
            RPC = None
    else:
        if RPC is None:
            RPC = Presence(client_id=client_id)
            RPC.connect()
            start_time = int(time.time())

        title = notion_window.title
        RPC.update(start=start_time, details=("Editing: " + title))

    time.sleep(15)