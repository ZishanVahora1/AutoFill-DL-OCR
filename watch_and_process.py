import os
import sys
import time
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

BASE_DIR = os.path.dirname(__file__)
DIRECTORY_TO_WATCH = BASE_DIR          # or os.path.join(BASE_DIR, "uploads")
PY_EXE = sys.executable

# Child output should show immediately in Terminal
ENV = os.environ.copy()
ENV["PYTHONUNBUFFERED"] = "1"

# Use service-account automatically if present
DLOCR = os.path.join(BASE_DIR, "DLOCR.json")
if os.path.exists(DLOCR):
    ENV["GOOGLE_APPLICATION_CREDENTIALS"] = DLOCR

# Reduce gRPC log spam
ENV.setdefault("GRPC_VERBOSITY", "ERROR")
ENV.setdefault("GRPC_TRACE", "")

class Handler(FileSystemEventHandler):
    def __init__(self):
        self._seen = {}  # (path, mtime) -> timestamp

    def on_created(self, event):
        if event.is_directory:
            return
        lower = event.src_path.lower()
        if not lower.endswith((".png", ".jpg", ".jpeg")):
            return

        # small delay so the copy finishes
        time.sleep(0.5)

        try:
            mtime = int(os.path.getmtime(event.src_path))
        except FileNotFoundError:
            return
        key = (event.src_path, mtime)
        if key in self._seen:
            return
        self._seen[key] = time.time()

        print(f"Processing file: {event.src_path}")
        subprocess.run(
            [PY_EXE, "-u", os.path.join(BASE_DIR, "detect_test.py"), event.src_path],  # -u = unbuffered
            env=ENV,
            check=False,
        )

def run():
    print(f"Watching: {DIRECTORY_TO_WATCH}")
    observer = Observer()
    observer.schedule(Handler(), DIRECTORY_TO_WATCH, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(5)
    except KeyboardInterrupt:
        observer.stop()
        print("Observer Stopped")
    observer.join()

if __name__ == "__main__":
    run()
