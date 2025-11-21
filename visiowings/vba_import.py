"""VBA Module Import functionality
Document module overwrite logic (force option)
Removes VBA header when importing via force
Supports multiple documents (drawings + stencils)
Automatically repairs missing VBA headers for new .bas files (VS Code workflow)
Improved: Robust header repair and user safety for imports
"""
import win32com.client
import pythoncom
from pathlib import Path
import re
from .document_manager import VisioDocumentManager

class VisioVBAImporter:
    def __init__(self, visio_file_path, force_document=False, debug=False, silent_reconnect=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.force_document = force_document
        self.debug = debug
        self.silent_reconnect = silent_reconnect
        self.doc_manager = None
        self.document_map = {}

    def connect_to_visio(self):
        try:
            try: pythoncom.CoInitialize()
            except: pass
            self.doc_manager = VisioDocumentManager(self.visio_file_path, debug=self.debug)
            if not self.doc_manager.connect_to_visio():
                return False  
            self.visio_app = self.doc_manager.visio_app
            self.doc = self.doc_manager.main_doc
            for doc_info in self.doc_manager.get_all_documents_with_vba():
                self.document_map[doc_info.folder_name] = doc_info
            if self.debug:
                print(f"[DEBUG] Document map created: {list(self.document_map.keys())}")
            return True
        except Exception as e:
            print(f"❌ Connection error: {e}")
            if self.debug:
                import traceback; traceback.print_exc()
            self.doc = None
            return False
    def _ensure_connection(self):
        try:
            _ = self.doc.Name
            return True
        except Exception:
            if self.debug and not self.silent_reconnect:
                print("[DEBUG] Connection lost, attempting to reconnect...")
            elif not self.debug and not self.silent_reconnect:
                print("🔄 Connection lost, attempting to reconnect...")
            return self.connect_to_visio()
    def _strip_vba_header(self, code):
        lines = code.splitlines()
        filtered_lines = []
        started = False
        for line in lines:
            # Remove VERSION, Begin, End, MultiUse lines
            if line.startswith(('VERSION', 'Begin', 'End', 'MultiUse')):
                continue
            # Remove Attribute lines except VB_Name
            if line.startswith('Attribute '):
                if 'VB_Name' in line:
                    filtered_lines.append(line)
                continue
            # Preserve comments anywhere in file
            if line.strip().startswith("'"):
                filtered_lines.append(line)
                continue
            # Once reaching Option Explicit or any non-header, start keeping lines
            if line.strip() and not started:
                started = True
            if started:
                filtered_lines.append(line)
        # Remove leading/trailing blanks
        while filtered_lines and not filtered_lines[0].strip():
            filtered_lines.pop(0)
        while filtered_lines and not filtered_lines[-1].strip():
            filtered_lines.pop()
        return '\n'.join(filtered_lines)
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
    def _repair_vba_module_file(self, file_path):
        text = ""
        try:
            text = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            try:
                text = file_path.read_text(encoding="cp1252")
            except Exception:
                text = ""
        except Exception:
            text = ""
        module_name = file_path.stem
        header = f'Attribute VB_Name = "{module_name}"\nOption Explicit\n'
        dummy_sub = f'Sub Dummy()\nEnd Sub\n'
        needs_write = False
        has_header = 'Attribute VB_Name' in text
        # Fix header with proper logic
        stripped = self._strip_vba_header(text)
        if not stripped.strip():
            text = header + '\n' + dummy_sub
            needs_write = True
        elif not has_header:
            text = header + stripped
            needs_write = True
        else:
            text = stripped
            needs_write = True
        if text and not text.endswith("\n"):
            text += "\n"
            needs_write = True
        if needs_write:
            try:
                file_path.write_text(text, encoding="cp1252", errors='replace')
            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] Error writing as cp1252, trying UTF-8: {e}")
                file_path.write_text(text, encoding="utf-8")
    def import_module(self, file_path):
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print(f"[DEBUG] COM initialized for import_module thread")
        except:
            if self.debug:
                print(f"[DEBUG] COM already initialized in this thread")
            pass
        try:
            if not self._ensure_connection():
                print("⚠️  No connection to Visio - make sure the document is open")
                return False
            file_path = Path(file_path)
            self._repair_vba_module_file(file_path)
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
                    except UnicodeDecodeError:
                        code = file_path.read_text(encoding="cp1252", errors='replace')
                    code = self._strip_vba_header(code)
                    cm = component.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if code.strip():
                        cm.AddFromString(code)
                    try:
                        new_text = file_path.read_text(encoding="cp1252", errors='replace')
                        try:
                            current_text = file_path.read_text(encoding="utf-8")
                            if new_text != current_text:
                                file_path.write_text(new_text, encoding="utf-8")
                                if self.debug:
                                    print(f"[DEBUG] Converted {file_path.name} back to UTF-8 (content changed)")
                            elif self.debug:
                                print(f"[DEBUG] Skipped conversion - content unchanged")
                        except:
                            file_path.write_text(new_text, encoding="utf-8")
                            if self.debug:
                                print(f"[DEBUG] Converted {file_path.name} to UTF-8")
                    except Exception as e:
                        if self.debug:
                            print(f"[DEBUG] Error converting back to UTF-8: {e}")
                    print(f"✓ Imported: {target_doc_info.folder_name}/{file_path.name} (force)")
                    return True
                else:
                    print(f"⚠️  Document module '{module_name}' skipped without --force.")
                    if self.debug: print("[DEBUG] Use --force to overwrite document modules")
                    return False
            if component:
                if self.debug: print(f"[DEBUG] Removing existing module: {module_name}")
                vb_project.VBComponents.Remove(component)
            vb_project.VBComponents.Import(str(file_path))
            try:
                new_text = file_path.read_text(encoding="cp1252", errors='replace')
                try:
                    current_text = file_path.read_text(encoding="utf-8")
                    if new_text != current_text:
                        file_path.write_text(new_text, encoding="utf-8")
                        if self.debug:
                            print(f"[DEBUG] Converted {file_path.name} back to UTF-8 (content changed)")
                    elif self.debug:
                        print(f"[DEBUG] Skipped conversion - content unchanged")
                except:
                    file_path.write_text(new_text, encoding="utf-8")
                    if self.debug:
                        print(f"[DEBUG] Converted {file_path.name} to UTF-8")
            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] Error converting back to UTF-8: {e}")
            print(f"✓ Imported: {target_doc_info.folder_name}/{file_path.name}")
            return True
        except Exception as e:
            print(f"✗ Error importing {file_path.name}: {e}")
            if self.debug:
                import traceback; traceback.print_exc()
            return False
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print(f"[DEBUG] COM uninitialized for import_module thread")
                except:
                    pass
    def get_document_folders(self):
        return list(self.document_map.keys())
