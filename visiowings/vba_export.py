"""VBA Module Export functionality: improved header stripping and hash-based change detection"""
import win32com.client
import os
from pathlib import Path
import re
import hashlib

class VisioVBAExporter:
    def __init__(self, visio_file_path, debug=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.debug = debug
    
    def connect_to_visio(self):
        """Connect to Visio application and open document"""
        try:
            self.visio_app = win32com.client.Dispatch("Visio.Application")
            # Try to find already open document first
            for doc in self.visio_app.Documents:
                if doc.FullName.lower() == str(self.visio_file_path).lower():
                    self.doc = doc
                    if self.debug:
                        print(f"[DEBUG] Verbunden mit geöffnetem Dokument: {doc.Name}")
                    return True
            
            # If not open, open it
            self.doc = self.visio_app.Documents.Open(self.visio_file_path)
            if self.debug:
                print(f"[DEBUG] Dokument geöffnet: {self.doc.Name}")
            return True
        except Exception as e:
            print(f"❌ Fehler beim Verbinden mit Visio: {e}")
            return False
    
    def _strip_vba_header_file(self, file_path):
        """Remove VBA header from exported file"""
        try:
            text = Path(file_path).read_text(encoding="utf-8")
            
            # More comprehensive header markers
            header_markers = [
                'Option Explicit',
                'Option Compare',
                'Option Base',
                'Sub ',
                'Function ',
                'Property ',
                'Public ',
                'Private ',
                'Dim '
            ]
            
            lines = text.splitlines()
            code_start = 0
            
            # Find first line that contains actual code
            for i, line in enumerate(lines):
                stripped = line.strip()
                # Skip empty lines and VBA metadata
                if not stripped or stripped.startswith(('VERSION', 'Begin', 'End', 'Attribute ', "'")):
                    continue
                # Check if this is actual code
                if any(marker in line for marker in header_markers):
                    code_start = i
                    break
            
            new_text = '\n'.join(lines[code_start:])
            Path(file_path).write_text(new_text, encoding="utf-8")
            
            if self.debug:
                removed_lines = code_start
                if removed_lines > 0:
                    print(f"[DEBUG] {removed_lines} Header-Zeilen entfernt aus {file_path.name}")
        
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Fehler beim Header-Stripping: {e}")
            pass
    
    def _module_content_hash(self, vb_project):
        """Generate content hash of all modules for change detection"""
        try:
            code_parts = []
            for comp in vb_project.VBComponents:
                cm = comp.CodeModule
                # Only hash actual code, not headers
                if cm.CountOfLines > 0:
                    code = cm.Lines(1, cm.CountOfLines)
                    code_parts.append(f"{comp.Name}:{code}")
            
            hash_input = ''.join(code_parts)
            content_hash = hashlib.md5(hash_input.encode()).hexdigest()
            
            if self.debug:
                print(f"[DEBUG] Hash berechnet: {content_hash[:8]}... ({len(code_parts)} Module)")
            
            return content_hash
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Fehler bei Hash-Berechnung: {e}")
            return None
    
    def export_modules(self, output_dir, last_hash=None):
        """Export VBA modules, only if content changed (via hash comparison)
        
        Returns:
            tuple: (list of exported files, current hash)
                  Returns ([], last_hash) if no changes detected
        """
        if not self.doc:
            print("❌ Kein Dokument geöffnet")
            return [], None
        
        try:
            vb_project = self.doc.VBProject
            
            # Calculate current hash
            current_hash = self._module_content_hash(vb_project)
            
            if self.debug:
                print(f"[DEBUG] Last hash: {last_hash[:8] if last_hash else 'None'}...")
                print(f"[DEBUG] Current hash: {current_hash[:8] if current_hash else 'None'}...")
            
            # Check if content actually changed
            if last_hash and last_hash == current_hash:
                if self.debug:
                    print("[DEBUG] Hashes identisch - kein Export")
                else:
                    print("✓ Keine Änderungen erkannt – kein Export notwendig")
                return [], current_hash  # Return empty list but current hash
            
            # Content changed, perform export
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            exported_files = []
            
            for component in vb_project.VBComponents:
                # Map component types to file extensions
                ext_map = {
                    1: '.bas',    # vbext_ct_StdModule
                    2: '.cls',    # vbext_ct_ClassModule
                    3: '.frm',    # vbext_ct_MSForm
                    100: '.cls'   # vbext_ct_Document
                }
                
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = output_path / file_name
                
                # Export module
                component.Export(str(file_path))
                
                # Remove VBA header for standard modules and class modules
                if component.Type in [1, 2, 100]:
                    self._strip_vba_header_file(file_path)
                
                exported_files.append(file_path)
                print(f"✓ Exportiert: {file_name}")
            
            return exported_files, current_hash
        
        except Exception as e:
            print(f"❌ Fehler beim Exportieren: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            else:
                print("")
                print("⚠️  Stelle sicher, dass in Visio folgende Einstellung aktiviert ist:")
                print("   Datei → Optionen → Trust Center → Trust Center-Einstellungen")
                print("   → Makroeinstellungen → 'Zugriff auf VBA-Projektobjektmodell vertrauen'")
            
            return [], None
