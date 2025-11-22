import win32com.client
import pythoncom
from pathlib import Path
import re
import sys
import os
from .document_manager import VisioDocumentManager
from difflib import unified_diff

class VisioVBAImporter:
    def __init__(self, visio_file_path, force_document=False, debug=False, silent_reconnect=False, always_yes=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.force_document = force_document
        self.debug = debug
        self.silent_reconnect = silent_reconnect
        self.doc_manager = None
        self.document_map = {}
        self.always_yes = always_yes

    def connect_to_visio(self):
        try:
            pythoncom.CoInitialize()
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] COM already initialized: {e}")
        self.doc_manager = VisioDocumentManager(self.visio_file_path, debug=self.debug)
        if not self.doc_manager.connect_to_visio():
            return False
        self.visio_app = self.doc_manager.visio_app
        self.doc = self.doc_manager.main_doc
        if not self.doc:
            print("❌ Failed to connect to main document")
            return False
        for doc_info in self.doc_manager.get_all_documents_with_vba():
            self.document_map[doc_info.folder_name] = doc_info
        if self.debug:
            print(f"[DEBUG] Document map created: {list(self.document_map.keys())}")
        return True

    def _ensure_connection(self):
        try:
            _ = self.doc.Name
            return True
        except Exception as e:
            if self.debug and not self.silent_reconnect:
                print(f"[DEBUG] Connection lost ({e}), attempting to reconnect...")
            elif not self.debug and not self.silent_reconnect:
                print("🔄 Connection lost, attempting to reconnect...")
            return self.connect_to_visio()

    def _find_document_for_file(self, file_path):
        parent_dir = file_path.parent.name
        if parent_dir in self.document_map:
            if self.debug:
                print(f"[DEBUG] File {file_path.name} belongs to document: {parent_dir}")
            return self.document_map[parent_dir]
        main_doc_info = self.doc_manager.get_main_document()
        if self.debug:
            print(f"[DEBUG] File {file_path.name} assigned to main document")
        return main_doc_info

    def _create_temp_cp1252_file(self, file_path):
        import tempfile
        try:
            text = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            text = file_path.read_text(encoding="cp1252")
        module_name = file_path.stem
        if "Attribute VB_Name" not in text:
            header = f'Attribute VB_Name = "{module_name}"\n'
            text = header + text
        if text and not text.endswith("\n"):
            text += "\n"
        fd, temp_path = tempfile.mkstemp(suffix=file_path.suffix, text=True)
        try:
            with os.fdopen(fd, 'w', encoding='cp1252') as f:
                f.write(text)
            if self.debug:
                print(f"[DEBUG] Created temp CP1252 file: {temp_path}")
            return temp_path
        except UnicodeEncodeError as e:
            print(f"⚠️  Warning: {file_path.name} contains characters not supported in CP1252")
            with os.fdopen(fd, 'w', encoding='cp1252', errors='replace') as f:
                f.write(text)
            return temp_path
        except Exception:
            os.close(fd)
            os.unlink(temp_path)
            raise

    def import_module(self, file_path):
        com_initialized = False
        temp_file = None
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print(f"[DEBUG] COM initialized for import_module thread")
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] COM already initialized in this thread: {e}")
        try:
            if not self.connect_to_visio():
                print("⚠️  No connection to Visio - make sure the document is open")
                return False
            file_path = Path(file_path)
            target_doc_info = self._find_document_for_file(file_path)
            if not target_doc_info:
                print(f"⚠️  No matching document found for {file_path.name}")
                return False
            vb_project = target_doc_info.doc.VBProject
            module_name = file_path.stem
            if self.debug:
                print(f"[DEBUG] Importing {file_path.name} into {target_doc_info.name}")
            component = None
            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    component = comp
                    break
            if component and component.Type == 100:
                if self.force_document:
                    try:
                        code = file_path.read_text(encoding="utf-8")
                    except Exception:
                        code = file_path.read_text(encoding="cp1252", errors='replace')
                    code = self._strip_vba_header(code)
                    cm = component.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if code.strip():
                        cm.AddFromString(code)
                    print(f"✓ Imported: {target_doc_info.folder_name}/{file_path.name} (force)")
                    return True
                else:
                    print(f"⚠️  Document module '{module_name}' skipped without --force.")
                    return False
            if component:
                if not self._prompt_overwrite(module_name, file_path, component):
                    print(f"⊘ Skipped: {module_name}")
                    return False
                vb_project.VBComponents.Remove(component)
            temp_file = self._create_temp_cp1252_file(file_path)
            vb_project.VBComponents.Import(str(temp_file))
            print(f"✓ Imported: {target_doc_info.folder_name}/{file_path.name}")
            return True
        except Exception as e:
            print(f"✗ Error importing {file_path}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return False
        finally:
            if temp_file and temp_file != str(file_path):
                try:
                    os.unlink(temp_file)
                    if self.debug:
                        print(f"[DEBUG] Cleaned up temp file: {temp_file}")
                except Exception as e:
                    if self.debug:
                        print(f"[DEBUG] Error cleaning temp file: {e}")
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print(f"[DEBUG] COM uninitialized for import_module thread")
                except Exception as e:
                    if self.debug:
                        print(f"[DEBUG] Error uninitializing COM: {e}")

    # ... keep other methods unchanged ...
