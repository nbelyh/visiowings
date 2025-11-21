"""VBA Module Export functionality: improved header stripping and hash-based change detection
Now supports multiple documents (drawings + stencils)
Enhanced: Remove local module files if deleted in Visio
"""
import win32com.client
import os
from pathlib import Path
import re
import hashlib
from .document_manager import VisioDocumentManager, VisioDocumentInfo

class VisioVBAExporter:
    def __init__(self, visio_file_path, debug=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.debug = debug
        self.doc_manager = None
    
    def connect_to_visio(self, silent=False):
        try:
            self.doc_manager = VisioDocumentManager(self.visio_file_path, debug=self.debug)
            if not self.doc_manager.connect_to_visio():
                return False
            self.visio_app = self.doc_manager.visio_app
            self.doc = self.doc_manager.main_doc
            if not silent:
                self.doc_manager.print_summary()
            return True
        except Exception as e:
            if not silent:
                print(f"❌ Error connecting to Visio: {e}")
            return False
    
    def _strip_vba_header_file(self, file_path):
        try:
            # Visio exports files in Windows-1252 (ANSI), so read with that encoding
            text = Path(file_path).read_text(encoding="cp1252")
            
            lines = text.splitlines()
            filtered_lines = []
            
            # Pattern to match VBA header lines to REMOVE (but keep Attribute VB_Name)
            # Remove: VERSION, Begin, End, MultiUse, and Attribute lines EXCEPT Attribute VB_Name
            for line in lines:
                stripped = line.strip()
                
                # Skip empty lines at the start
                if not stripped and not filtered_lines:
                    continue
                    
                # Skip VERSION, Begin, End, MultiUse
                if line.startswith(('VERSION', 'Begin', 'End', 'MultiUse')):
                    continue
                    
                # Skip Attribute lines EXCEPT Attribute VB_Name
                if line.startswith('Attribute '):
                    if 'VB_Name' not in line:
                        continue
                
                # Keep this line
                filtered_lines.append(line)
            
            new_text = '\n'.join(filtered_lines)
            
            # Write as UTF-8 for VS Code compatibility (converts from cp1252)
            Path(file_path).write_text(new_text, encoding="utf-8")
            
            if self.debug:
                removed_count = len(lines) - len(filtered_lines)
                if removed_count > 0:
                    print(f"[DEBUG] {removed_count} header lines removed from {file_path.name}")
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error during header stripping: {e}")
            pass
    
    def _module_content_hash(self, vb_project):
        try:
            code_parts = []
            for comp in vb_project.VBComponents:
                cm = comp.CodeModule
                if cm.CountOfLines > 0:
                    code = cm.Lines(1, cm.CountOfLines)
                    code_parts.append(f"{comp.Name}:{code}")
            hash_input = ''.join(code_parts)
            content_hash = hashlib.md5(hash_input.encode()).hexdigest()
            if self.debug:
                print(f"[DEBUG] Hash calculated: {content_hash[:8]}... ({len(code_parts)} modules)")
            return content_hash
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error during hash calculation: {e}")
            return None
    
    def _compare_files(self, local_path, temp_export_path):
        """Compare local file with newly exported file"""
        try:
            local_content = local_path.read_text(encoding="utf-8")
            export_content = temp_export_path.read_text(encoding="cp1252")
            return local_content == export_content
        except Exception:
            return False
    
    def _export_document_modules(self, doc_info, output_dir, last_hash=None):
        try:
            vb_project = doc_info.doc.VBProject
            current_hash = self._module_content_hash(vb_project)
            
            if last_hash and last_hash == current_hash:
                if self.debug:
                    print(f"[DEBUG] {doc_info.name}: Hashes identical - no export")
                # Even if no export, check if Visio modules were deleted!
                self._sync_deleted_modules(doc_info, output_dir, vb_project)
                return [], current_hash
            
            doc_output_path = Path(output_dir) / doc_info.folder_name
            doc_output_path.mkdir(parents=True, exist_ok=True)
            
            # Check for files with local changes
            files_with_changes = []
            ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}
            
            for component in vb_project.VBComponents:
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = doc_output_path / file_name
                
                # If file exists locally, check if it would be different
                if file_path.exists() and component.Type in [1, 2, 100]:
                    # Export to temp location to compare
                    import tempfile
                    with tempfile.NamedTemporaryFile(mode='w', suffix=ext, delete=False, encoding='utf-8') as tmp:
                        temp_path = Path(tmp.name)
                    
                    try:
                        component.Export(str(temp_path))
                        
                        # Process temp file the same way we would during export
                        self._strip_vba_header_file(temp_path)
                        
                        # Now compare processed files (both UTF-8)
                        try:
                            local_content = file_path.read_text(encoding="utf-8")
                            temp_content = temp_path.read_text(encoding="utf-8")
                            
                            if local_content != temp_content:
                                files_with_changes.append(file_name)
                        except Exception as e:
                            if self.debug:
                                print(f"[DEBUG] Error comparing {file_name}: {e}")
                    finally:
                        temp_path.unlink(missing_ok=True)
            
            # If there are files with local changes, ask user
            if files_with_changes:
                print(f"\n⚠️  The following local files differ from Visio:")
                for fname in files_with_changes:
                    print(f"   - {doc_info.folder_name}/{fname}")
                
                response = input(f"\nOverwrite local files with Visio content? (y/N): ").strip().lower()
                if response not in ('y', 'yes'):
                    print(f"❌ Export cancelled for {doc_info.name}")
                    return [], None
            
            # Proceed with export
            exported_files = []
            visio_module_names = set()
            
            for component in vb_project.VBComponents:
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = doc_output_path / file_name
                component.Export(str(file_path))
                if component.Type in [1, 2, 100]:
                    self._strip_vba_header_file(file_path)
                exported_files.append(file_path)
                visio_module_names.add(component.Name.lower())
                print(f"✓ Exported: {doc_info.folder_name}/{file_name}")
            
            # After export, sync deleted local files
            self._sync_deleted_modules(doc_info, output_dir, vb_project, visio_module_names)
            
            return exported_files, current_hash
        except Exception as e:
            print(f"❌ Error exporting {doc_info.name}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return [], None
    
    def _sync_deleted_modules(self, doc_info, output_dir, vb_project, visio_module_names=None):
        doc_output_path = Path(output_dir) / doc_info.folder_name
        local_files = list(doc_output_path.glob("*.bas")) + list(doc_output_path.glob("*.cls")) + list(doc_output_path.glob("*.frm"))
        
        if visio_module_names is None:
            visio_module_names = set(comp.Name.lower() for comp in vb_project.VBComponents)
        
        # Collect files to delete
        files_to_delete = []
        for file in local_files:
            filename = file.stem.lower()
            if filename not in visio_module_names:
                files_to_delete.append(file)
        
        # If there are files to delete, ask user
        if files_to_delete:
            print(f"\n⚠️  The following local files are missing in Visio:")
            for file in files_to_delete:
                print(f"   - {doc_info.folder_name}/{file.name}")
            
            print(f"\nOptions:")
            print(f"  d - Delete local files")
            print(f"  i - Import to Visio")
            print(f"  k - Keep local files (default)")
            response = input(f"\nChoose action (d/i/K): ").strip().lower()
            
            if response == 'd':
                for file in files_to_delete:
                    try:
                        file.unlink()
                        print(f"✓ Removed local file: {doc_info.folder_name}/{file.name}")
                    except Exception as e:
                        print(f"⚠️  Could not remove local file: {file} ({e})")
            elif response == 'i':
                print(f"\n📤 Importing {len(files_to_delete)} file(s) to Visio...")
                for file in files_to_delete:
                    try:
                        vb_project.VBComponents.Import(str(file))
                        print(f"✓ Imported to Visio: {doc_info.folder_name}/{file.name}")
                    except Exception as e:
                        print(f"✗ Error importing {file.name}: {e}")
            else:
                print(f"ℹ️  Kept {len(files_to_delete)} local file(s)")
    
    def export_modules(self, output_dir, last_hashes=None):
        if not self.doc_manager:
            print("❌ No document manager initialized")
            return {}, {}
        
        if last_hashes is None:
            last_hashes = {}
        
        all_exported = {}
        all_hashes = {}
        
        documents = self.doc_manager.get_all_documents_with_vba()
        if not documents:
            print("⚠️  No documents with VBA code found")
            return {}, {}
        
        try:
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            for doc_info in documents:
                if self.debug:
                    print(f"[DEBUG] Exporting {doc_info.name}...")
                
                last_hash = last_hashes.get(doc_info.folder_name)
                exported_files, current_hash = self._export_document_modules(
                    doc_info, 
                    output_dir, 
                    last_hash
                )
                
                if exported_files or current_hash:
                    all_exported[doc_info.folder_name] = exported_files
                    all_hashes[doc_info.folder_name] = current_hash
            
            return all_exported, all_hashes
        except Exception as e:
            print(f"❌ Error during export: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            else:
                print("")
                print("⚠️  Make sure the following setting is enabled in Visio:")
                print("   File → Options → Trust Center → Trust Center Settings")
                print("   → Macro Settings → 'Trust access to the VBA project object model'")
            return {}, {}
