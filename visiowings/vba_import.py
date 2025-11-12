"""VBA Module Import functionality
Document module overwrite logic (force option)
Removes VBA header when importing via force
Supports multiple documents (drawings + stencils)
"""
import win32com.client
from pathlib import Path
import re
from .document_manager import VisioDocumentManager

class VisioVBAImporter:
    """Importiert VBA-Module in Visio-Dokumente, optional mit force für Document-Module"""
    
    def __init__(self, visio_file_path, force_document=False, debug=False, silent_reconnect=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.force_document = force_document
        self.debug = debug
        self.silent_reconnect = silent_reconnect  # Neu: unterdrückt Output bei expected reconnects
        self.doc_manager = None
        self.document_map = {}  # Maps folder_name -> VisioDocumentInfo

    def connect_to_visio(self):
        """Verbindet sich mit bereits geöffnetem Dokument und entdeckt alle Dokumente"""
        try:
            self.doc_manager = VisioDocumentManager(self.visio_file_path, debug=self.debug)
            if not self.doc_manager.connect_to_visio():
                return False            
            self.visio_app = self.doc_manager.visio_app
            self.doc = self.doc_manager.main_doc
            # Build document map (folder_name -> document)
            for doc_info in self.doc_manager.get_all_documents_with_vba():
                self.document_map[doc_info.folder_name] = doc_info
            if self.debug:
                print(f"[DEBUG] Dokument-Map erstellt: {list(self.document_map.keys())}")
            return True
        except Exception as e:
            print(f"❌ Fehler beim Verbinden: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self.doc = None
            return False

    def _ensure_connection(self):
        """Stellt sicher, dass die Verbindung zum Dokument noch aktiv ist"""
        try:
            _ = self.doc.Name
            return True
        except Exception:
            # Nur Nachricht wenn NICHT silent_reconnect
            if self.debug and not self.silent_reconnect:
                print("[DEBUG] Verbindung verloren, versuche neu zu verbinden...")
            elif not self.debug and not self.silent_reconnect:
                print("🔄 Verbindung verloren, versuche neu zu verbinden...")
            return self.connect_to_visio()
    
    def _strip_vba_header(self, code):
        lines = code.splitlines()
        code_start = 0
        vba_header_pattern = re.compile(r'^(VERSION|Begin|End|Attribute |MultiUse)')
        for i, line in enumerate(lines):
            stripped = line.strip()
            if not stripped or stripped.startswith("'"):
                continue
            if vba_header_pattern.match(line):
                continue
            code_start = i
            break
        result = '\n'.join(lines[code_start:])
        if self.debug and code_start > 0:
            print(f"[DEBUG] {code_start} Header-Zeilen beim Import entfernt")
        return result

    def _find_document_for_file(self, file_path):
        parent_dir = file_path.parent.name
        if parent_dir in self.document_map:
            if self.debug:
                print(f"[DEBUG] Datei {file_path.name} gehört zu Dokument: {parent_dir}")
            return self.document_map[parent_dir]
        main_doc_info = self.doc_manager.get_main_document()
        if self.debug:
            print(f"[DEBUG] Datei {file_path.name} wird Hauptdokument zugeordnet")
        return main_doc_info

    def import_module(self, file_path):
        # Verbindung vor jedem Import prüfen
        if not self._ensure_connection():
            print("⚠️  Keine Verbindung zu Visio - stelle sicher, dass das Dokument geöffnet ist")
            return False
        try:
            file_path = Path(file_path)
            target_doc_info = self._find_document_for_file(file_path)
            if not target_doc_info:
                print(f"⚠️  Kein passendes Dokument für {file_path.name} gefunden")
                return False
            vb_project = target_doc_info.doc.VBProject
            module_name = file_path.stem
            if self.debug:
                print(f"[DEBUG] Importiere {file_path.name} in {target_doc_info.name}")
            component = None
            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    component = comp
                    break
            if component and component.Type == 100:
                if self.force_document:
                    code = file_path.read_text(encoding="utf-8")
                    code = self._strip_vba_header(code)
                    cm = component.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if code.strip():
                        cm.AddFromString(code)
                    print(f"✓ Importiert: {target_doc_info.folder_name}/{file_path.name} (force)")
                    return True
                else:
                    print(f"⚠️  Document-Module '{module_name}' ohne --force übersprungen.")
                    if self.debug:
                        print("[DEBUG] Verwende --force um Document-Module zu überschreiben")
                    return False
            if component:
                if self.debug:
                    print(f"[DEBUG] Entferne existierendes Modul: {module_name}")
                vb_project.VBComponents.Remove(component)
            vb_project.VBComponents.Import(str(file_path))
            print(f"✓ Importiert: {target_doc_info.folder_name}/{file_path.name}")
            return True
        except Exception as e:
            print(f"✗ Fehler beim Importieren von {file_path.name}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return False
    def get_document_folders(self):
        return list(self.document_map.keys())
