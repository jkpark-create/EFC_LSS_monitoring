import sys
sys.path.append(".")
import time
from xplatform_icc_helper import collect_windows, main_window, click_rel, bring_to_front
from pywinauto import keyboard

windows = collect_windows()
info = main_window(windows)

if info:
    print(f"Targeting window: {info.title}")
    bring_to_front(info)
    time.sleep(1)
    
    # Try clicking Document tab
    print("Clicking Document tab at 450, 55...")
    click_rel(info, 450, 55)
    time.sleep(1)
    
    # Try clicking Search box
    print("Clicking Search box at 1000, 55...")
    click_rel(info, 1000, 55, double=True)
    time.sleep(1)
    
    # Type something
    print("Typing 'TEST'...")
    keyboard.send_keys("TEST", with_spaces=True)
    time.sleep(1)
    
    from xplatform_icc_helper import capture_window
    from pathlib import Path
    output_dir = Path("logs/debug")
    capture = capture_window(info, output_dir, prefix="click_test")
    print(f"Captured click test screenshot: {capture.resolve()}")
else:
    print("Main window not found.")
