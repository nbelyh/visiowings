"""VBA Module Export functionality"""

import win32com.client
import os
from pathlib import Path


class VisioVBAExporter:
    """Exportiert VBA-Module aus Visio-Dokumenten"""
    
    def __init__(self, visio_file_path):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        
    def connect_to_visio(self):
        """Verbindet sich mit Visio und öffnet das Dokument"""
        try:
            self.visio_app = win32com.client.Dispatch("Visio.Application")
            self.doc = self.visio_app.Documents.Open(self.visio_file_path)
            return True
        except Exception as e:
            print(f"❌ Fehler beim Verbinden mit Visio: {e}")
            return False
    
    def export_modules(self, output_dir):
        """Exportiert alle VBA-Module in ein Verzeichnis"""
        if not self.doc:
            print("❌ Kein Dokument geöffnet")
            return []
        
        try:
            vb_project = self.doc.VBProject
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            exported_files = []
            
            for component in vb_project.VBComponents:
                # Bestimme Dateierweiterung basierend auf Typ
                # 1 = vbext_ct_StdModule
                # 2 = vbext_ct_ClassModule
                # 3 = vbext_ct_MSForm
                # 100 = vbext_ct_Document
                ext_map = {
                    1: '.bas',
                    2: '.cls',
                    3: '.frm',
                    100: '.cls'
                }
                
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = output_path / file_name
                
                # Exportiere das Modul
                component.Export(str(file_path))
                exported_files.append(file_path)
                print(f"✓ Exportiert: {file_name}")
            
            return exported_files
            
        except Exception as e:
            print(f"❌ Fehler beim Exportieren: {e}")
            print("")
            print("⚠️  Stelle sicher, dass in Visio folgende Einstellung aktiviert ist:")
            print("   Datei → Optionen → Trust Center → Trust Center-Einstellungen")
            print("   → Makroeinstellungen → 'Zugriff auf VBA-Projektobjektmodell vertrauen'")
            return []
