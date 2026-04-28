import sys
sys.path.append(".")
from pathlib import Path
from xplatform_icc_helper import collect_windows, print_windows, best_capture_window, capture_window

windows = collect_windows()
print(f"Found {len(windows)} XPlatform windows.")
print_windows(windows)

target = best_capture_window(windows)
if target:
    output_dir = Path("logs/debug")
    output_dir.mkdir(parents=True, exist_ok=True)
    capture = capture_window(target, output_dir, prefix="debug_check")
    print(f"Captured debug screenshot: {capture.resolve()}")
else:
    print("No visible XPlatform window to capture.")
