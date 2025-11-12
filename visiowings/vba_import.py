"""VBA Module Import functionality"""

import win32com.client
from pathlib import Path


class VisioVBAImporter:
    """Importiert VBA-Module in Visio-Dokumente"""
    
    def __init__(self, visio_file_path):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
    
    def connect_to_visio(self):
        """Verbindet sich mit bereits geöffnetem Dokument"""
        try:
            self.visio_app = win32com.client.Dispatch("Visio.Application")
            
            # Suche nach bereits geöffnetem Dokument
            for doc in self.visio_app.Documents:
                if doc.FullName.lower() == str(self.visio_file_path).lower():
                    self.doc = doc
                    return True
            
            print(f"⚠️  Dokument nicht geöffnet: {self.visio_file_path}")
            print("   Bitte öffne das Dokument in Visio.")
            return False
            
        except Exception as e:
            print(f"❌ Fehler beim Verbinden: {e}")
            return False
    
    def import_module(self, file_path):
        """Importiert ein einzelnes VBA-Modul"""
        if not self.doc:
            return False
        
        try:
            vb_project = self.doc.VBProject
            file_path = Path(file_path)
            module_name = file_path.stem
            
            # Entferne existierendes Modul falls vorhanden
            for component in vb_project.VBComponents:
                if component.Name == module_name:
                    vb_project.VBComponents.Remove(component)
                    break
            
            # Importiere neues Modul
            vb_project.VBComponents.Import(str(file_path))
            print(f"✓ Importiert: {file_path.name}")
            return True
            
        except Exception as e:
            print(f"✗ Fehler beim Importieren von {file_path.name}: {e}")
            return False
