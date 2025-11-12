"""Command Line Interface for visiowings with bidirectional sync support"""
import argparse
from pathlib import Path
from .vba_export import VisioVBAExporter
from .vba_import import VisioVBAImporter
from .file_watcher import VBAWatcher

def cmd_edit(args):
    """Edit command: Export + Watch + Import with live sync"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    debug = getattr(args, 'debug', False)
    
    if not visio_file.exists():
        print(f"❌ Datei nicht gefunden: {visio_file}")
        return
    
    print(f"📂 Visio-Datei: {visio_file}")
    print(f"📁 Export-Verzeichnis: {output_dir}")
    if debug:
        print("[DEBUG] Debug-Modus aktiviert")
    
    print("\n=== Exportiere VBA-Module ===")
    exporter = VisioVBAExporter(str(visio_file), debug=debug)
    if not exporter.connect_to_visio():
        return
    
    result = exporter.export_modules(output_dir)
    if not result or len(result) != 2:
        print("❌ Keine Module exportiert")
        return
    
    exported_files, initial_hash = result
    if not exported_files:
        print("❌ Keine Module exportiert")
        return
    
    print(f"\n✓ {len(exported_files)} Module exportiert")
    if debug:
        print(f"[DEBUG] Initial hash: {initial_hash[:8]}...")
    
    print("\n=== Starte Live-Synchronisation ===")
    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug)
    if not importer.connect_to_visio():
        return
    
    watcher = VBAWatcher(
        output_dir, 
        importer, 
        exporter=exporter, 
        bidirectional=getattr(args, 'bidirectional', False),
        debug=debug
    )
    watcher.start()

def cmd_export(args):
    """Export command: Export VBA modules only"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    debug = getattr(args, 'debug', False)
    
    exporter = VisioVBAExporter(str(visio_file), debug=debug)
    if exporter.connect_to_visio():
        result = exporter.export_modules(output_dir)
        if result and len(result) == 2:
            exported_files, hash_value = result
            if exported_files:
                print(f"\n✓ {len(exported_files)} Module exportiert")
                if debug:
                    print(f"[DEBUG] Hash: {hash_value[:8]}...")

def cmd_import(args):
    """Import command: Import VBA modules only"""
    visio_file = Path(args.file).resolve()
    input_dir = Path(args.input or '.').resolve()
    debug = getattr(args, 'debug', False)
    
    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug)
    if importer.connect_to_visio():
        imported_count = 0
        for ext in ['*.bas', '*.cls', '*.frm']:
            for file in input_dir.glob(ext):
                if importer.import_module(file):
                    imported_count += 1
        
        if imported_count > 0:
            print(f"\n✓ {imported_count} Module importiert")
        else:
            print("\n⚠️  Keine Module gefunden oder importiert")

def main():
    parser = argparse.ArgumentParser(
        description='visiowings - VBA Editor für Visio mit VS Code Integration',
        epilog='Beispiel: visiowings edit --file dokument.vsdm --force --bidirectional --debug'
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Verfügbare Befehle')
    
    # Edit command
    edit_parser = subparsers.add_parser(
        'edit', 
        help='VBA-Module bearbeiten mit Live-Sync (VS Code ↔ Visio)'
    )
    edit_parser.add_argument('--file', '-f', required=True, help='Visio-Datei (.vsdm)')
    edit_parser.add_argument('--output', '-o', help='Export-Verzeichnis (Standard: aktuelles Verzeichnis)')
    edit_parser.add_argument(
        '--force', 
        action='store_true', 
        help='Document-Module (überschreiben (ThisDocument.cls)'
    )
    edit_parser.add_argument(
        '--bidirectional', 
        action='store_true', 
        help='Bidirektionaler Sync: Änderungen in Visio automatisch nach VS Code exportieren'
    )
    edit_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug-Modus: Ausführliche Log-Ausgaben'
    )
    
    # Export command
    export_parser = subparsers.add_parser('export', help='VBA-Module exportieren (einmalig)')
    export_parser.add_argument('--file', '-f', required=True, help='Visio-Datei (.vsdm)')
    export_parser.add_argument('--output', '-o', help='Export-Verzeichnis (Standard: aktuelles Verzeichnis)')
    export_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug-Modus: Ausführliche Log-Ausgaben'
    )
    
    # Import command
    import_parser = subparsers.add_parser('import', help='VBA-Module importieren (einmalig)')
    import_parser.add_argument('--file', '-f', required=True, help='Visio-Datei (.vsdm)')
    import_parser.add_argument('--input', '-i', help='Import-Verzeichnis (Standard: aktuelles Verzeichnis)')
    import_parser.add_argument(
        '--force', 
        action='store_true', 
        help='Document-Module überschreiben (ThisDocument.cls)'
    )
    import_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug-Modus: Ausführliche Log-Ausgaben'
    )
    
    args = parser.parse_args()
    
    if args.command == 'edit':
        cmd_edit(args)
    elif args.command == 'export':
        cmd_export(args)
    elif args.command == 'import':
        cmd_import(args)
    else:
        parser.print_help()

if __name__ == '__main__':
    main()
