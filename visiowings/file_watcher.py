"""File system watcher for VBA modules"""

import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path


class VBAFileHandler(FileSystemEventHandler):
    """Überwacht VBA-Dateien auf Änderungen"""
    
    def __init__(self, importer, extensions=['.bas', '.cls', '.frm']):
        self.importer = importer
        self.extensions = extensions
        self.last_modified = {}
    
    def on_modified(self, event):
        """Wird aufgerufen wenn eine Datei geändert wird"""
        if event.is_directory:
            return
        
        file_path = Path(event.src_path)
        
        # Nur relevante Dateierweiterungen
        if file_path.suffix not in self.extensions:
            return
        
        # Verhindere Mehrfach-Trigger (Debounce)
        current_time = time.time()
        last_time = self.last_modified.get(str(file_path), 0)
        
        if current_time - last_time < 1.0:  # Debounce 1 Sekunde
            return
        
        self.last_modified[str(file_path)] = current_time
        
        print(f"\n📝 Änderung erkannt: {file_path.name}")
        self.importer.import_module(file_path)


class VBAWatcher:
    """Startet den File Watcher"""
    
    def __init__(self, watch_directory, importer):
        self.watch_directory = watch_directory
        self.importer = importer
        self.observer = None
    
    def start(self):
        """Startet die Überwachung"""
        event_handler = VBAFileHandler(self.importer)
        self.observer = Observer()
        self.observer.schedule(
            event_handler, 
            str(self.watch_directory), 
            recursive=False
        )
        self.observer.start()
        
        print(f"\n👁️  Überwache Verzeichnis: {self.watch_directory}")
        print("💾 Speichere Dateien in VS Code (Ctrl+S) um sie nach Visio zu synchronisieren")
        print("⏸️  Drücke Ctrl+C zum Beenden...\n")
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop()
    
    def stop(self):
        """Stoppt die Überwachung"""
        if self.observer:
            self.observer.stop()
            self.observer.join()
            print("\n✓ Überwachung beendet")
