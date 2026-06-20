import os
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from consolidator import process_specific_files

SOURCE_FOLDER = r"\\atgv1pfss03\Groups\Enterprise Products\BBN\HFF EOL\2026\New 2026\New 2026"

# =====================================================
# DEBOUNCE — waits until file stops changing
# Excel saves temp ~$ files repeatedly while open,
# so we wait 8s of silence before processing.
# =====================================================

DEBOUNCE_SECONDS = 8

_pending_files = {}          # { normalised_path: (original_path, last_modified_ts) }
_lock          = threading.Lock()

watcher_status = {
    "last_file":  "",
    "last_time":  "",
    "processing": False,
    "message":    "Watching for changes..."
}


def _normalise(path):
    """Case- and separator-insensitive path key."""
    return os.path.normcase(os.path.normpath(path))


def _debounce_worker():
    while True:
        time.sleep(2)
        now = time.time()

        with _lock:
            # Collect files whose debounce window has elapsed
            ready_norm = [
                norm for norm, (orig, ts) in _pending_files.items()
                if now - ts >= DEBOUNCE_SECONDS
            ]
            # Grab originals and remove from pending dict
            ready_originals = []
            for norm in ready_norm:
                orig, _ = _pending_files.pop(norm)
                ready_originals.append(orig)

        if not ready_originals:
            continue

        watcher_status["processing"] = True
        watcher_status["message"]    = "Consolidating updated files..."

        print(f"\n[WATCHER] Files ready: {[os.path.basename(p) for p in ready_originals]}")

        try:
            count = process_specific_files(ready_originals)

            watcher_status["processing"] = False
            watcher_status["last_time"]  = time.strftime("%H:%M:%S")
            watcher_status["message"]    = (
                f"Auto-updated at {watcher_status['last_time']} "
                f"({count} DB file(s) updated)"
            )
            print(f"[WATCHER] Done — {count} DB file(s) updated.")

        except Exception as e:
            watcher_status["processing"] = False
            watcher_status["message"]    = f"Error: {str(e)}"
            print(f"[WATCHER] ERROR: {e}")


# =====================================================
# FILE EVENT HANDLER
# =====================================================

class ExcelHandler(FileSystemEventHandler):

    def _handle(self, path):
        filename = os.path.basename(path)

        if filename.startswith("~$"):
            return
        if not path.lower().endswith((".xlsx", ".xlsm")):
            return

        print(f"[WATCHER] Change detected: {filename}")
        watcher_status["last_file"] = filename
        watcher_status["message"]   = (
            f"Change detected in {filename}, waiting {DEBOUNCE_SECONDS}s..."
        )

        # FIX: store by normalised key so the same file doesn't
        #      create duplicate pending entries when events fire
        #      with slightly different path capitalisation.
        norm = _normalise(path)
        with _lock:
            _pending_files[norm] = (path, time.time())

    def on_modified(self, event):
        if not event.is_directory:
            self._handle(event.src_path)

    def on_created(self, event):
        if not event.is_directory:
            self._handle(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self._handle(event.dest_path)


# =====================================================
# START WATCHER
# =====================================================

def start_watcher():
    worker = threading.Thread(target=_debounce_worker, daemon=True)
    worker.start()

    handler  = ExcelHandler()
    observer = Observer()
    observer.schedule(handler, path=SOURCE_FOLDER, recursive=True)
    observer.daemon = True
    observer.start()

    print(f"[WATCHER] Watching : {SOURCE_FOLDER}")
    print(f"[WATCHER] Debounce : {DEBOUNCE_SECONDS} seconds")
    return observer


if __name__ == "__main__":
    obs = start_watcher()
    print("Watching... Press Ctrl+C to stop.\n")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()