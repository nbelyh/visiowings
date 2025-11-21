"""
VBA Module Import functionality
- Document module overwrite logic (force option)
- Preserves Classes, Modules, Forms, and ThisDocument structure
- User prompt before overwriting differing modules
- Robust error handling and header repair
"""
import win32com.client
import pythoncom
from pathlib import Path
import re
import sys
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

    def _module_type_from_ext(self, filename):
        ext = Path(filename).suffix.lower()
        if ext == ".bas":
            return "module"
        elif ext == ".cls":
            return "class"
        elif ext == ".frm":
            return "form"
        return "unknown"

    def _repair_vba_module_file(self, file_path):
        try:
            text = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            text = file_path.read_text(encoding="cp1252")
        module_name = file_path.stem
        header = f'Attribute VB_Name = "{module_name}"\nOption Explicit\n'
        if "Attribute VB_Name" not in text:
            text = header + text
        if text and not text.endswith("\n"):
            text += "\n"
        file_path.write_text(text, encoding="cp1252", errors='replace')
        return True

    def _read_module_code(self, file_path):
        try:
            return file_path.read_text(encoding="utf-8")
        except Exception:
            try:
                return file_path.read_text(encoding="cp1252")
            except Exception:
                return ""

    def _prompt_overwrite(self, module_name, file_path, comp):
        file_code = self._read_module_code(file_path)
        visio_code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines) if comp.CodeModule.CountOfLines > 0 else ""
        if file_code.strip() == visio_code.strip() or self.always_yes:
            return True
        print(f"\n⚠️  Module '{module_name}' differs from Visio. See diff below:")
        for line in unified_diff(visio_code.splitlines(), file_code.splitlines(), fromfile='Visio', tofile='Disk', lineterm=''):
            print(line)
        print(f"Overwrite module '{module_name}' in Visio with disk version? (y/N): ", end="")
        ans = input().strip().lower()
        return ans in ("y", "yes")

    def import_directory(self, input_dir):
        input_dir = Path(input_dir)
        # Structure: .../document_folder/Classes|Modules|Forms|VisioObjects/
        dirs = [d for d in input_dir.iterdir() if d.is_dir()]
        if not dirs:
            # backward compat: single-document root
            dirs = [input_dir]
        for doc_dir in dirs:
            # Structure inside doc_dir
            structure = {
                "Modules": [],
                "Classes": [],
                "Forms": [],
                "VisioObjects": [],
                "root": []
            }
            for subdir in doc_dir.iterdir():
                if not subdir.is_dir():
                    continue
                if subdir.name.lower() in ("modules", "classes", "forms", "visioobjects"):
                    structure[subdir.name.capitalize()].extend(subdir.glob("*.*"))
            # fallback: find files in root (legacy style)
            for f in doc_dir.glob("*.bas"):
                structure["Modules"].append(f)
            for f in doc_dir.glob("*.cls"):
                if f.parent.name.lower() != "visioobjects":
                    structure["Classes"].append(f)
                else:
                    structure["VisioObjects"].append(f)
            for f in doc_dir.glob("*.frm"):
                structure["Forms"].append(f)
            # Import order: Modules, Classes, Forms, VisioObjects (ThisDocument)
            for group in ("Modules", "Classes", "Forms", "VisioObjects"):
                for file_path in structure[group]:
                    module_type = self._module_type_from_ext(file_path)
                    module_name = file_path.stem
                    doc_info = self.document_map.get(doc_dir.name.lower())
                    if not doc_info:
                        print(f"❌ No document found for folder '{doc_dir.name}'")
                        continue
                    vb_project = doc_info.doc.VBProject
                    target_comp = None
                    for comp in vb_project.VBComponents:
                        if comp.Name == module_name:
                            target_comp = comp
                            break
                    # Optionally prompt user if overwriting
                    if target_comp is not None and group != "VisioObjects":
                        if not self._prompt_overwrite(module_name, file_path, target_comp):
                            print(f"⊘ Skipped: {module_name}")
                            continue
                        vb_project.VBComponents.Remove(target_comp)
                    # Always ask for ThisDocument unless --force_document
                    if group == "VisioObjects" and target_comp is not None and not self.force_document:
                        print(f"⚠️  Document module '{module_name}' skipped without --force.")
                        continue
                    self._repair_vba_module_file(file_path)
                    try:
                        vb_project.VBComponents.Import(str(file_path))
                        print(f"✓ Imported: {doc_dir.name}/{group}/{module_name}")
                    except Exception as e:
                        if group == "VisioObjects" and self.force_document and target_comp is not None:
                            # Overwrite code in place
                            with open(file_path, encoding='utf-8') as f:
                                code = f.read()
                            cm = target_comp.CodeModule
                            cm.DeleteLines(1, cm.CountOfLines)
                            cm.AddFromString(code)
                            print(f"✓ Imported: {doc_dir.name}/{group}/{module_name} (force)")
                        else:
                            print(f"❌ Failed to import {module_name}: {e}")
