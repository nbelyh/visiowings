"""Command Line Interface for visiowings"""

import argparse
from pathlib import Path
from .vba_export import VisioVBAExporter
from .vba_import import VisioVBAImporter
from .file_watcher import VBAWatcher


def cmd_edit(args):
    """Hauptfunktion für 'visiowings edit'"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    
    if not visio_file.exists():
        print(f"❌ Datei nicht gefunden: {visio_file}")
        return
    
    print(f"📂 Visio-Datei: {visio_file}")
    print(f"📁 Export-Verzeichnis: {output_dir}")
    
    # Phase 1: Exportieren
    print("\n=== Exportiere VBA-Module ===")
    exporter = VisioVBAExporter(str(visio_file))
    
    if not exporter.connect_to_visio():
        return
    
    exported_files = exporter.export_modules(output_dir)
    
    if not exported_files:
        print("❌ Keine Module exportiert")
        return
    
    print(f"\n✓ {len(exported_files)} Module exportiert")
    
    # Phase 2: Live-Sync starten
    print("\n=== Starte Live-Synchronisation ===")
    importer = VisioVBAImporter(str(visio_file))
    
    if not importer.connect_to_visio():
        return
    
    watcher = VBAWatcher(output_dir, importer)
    watcher.start()


def cmd_export(args):
    """Nur Export, kein Sync"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    
    exporter = VisioVBAExporter(str(visio_file))
    if exporter.connect_to_visio():
        exporter.export_modules(output_dir)


def cmd_import(args):
    """Nur Import, kein Sync"""
    visio_file = Path(args.file).resolve()
    input_dir = Path(args.input or '.').resolve()
    
    importer = VisioVBAImporter(str(visio_file))
    if importer.connect_to_visio():
        for ext in ['*.bas', '*.cls', '*.frm']:
            for file in input_dir.glob(ext):
                importer.import_module(file)


def main():
    """Haupteinstiegspunkt für CLI"""
    parser = argparse.ArgumentParser(
        description='visiowings - VBA Editor für Visio',
        epilog='Beispiel: visiowings edit --file dokument.vsdm'
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Verfügbare Befehle')
    
    # edit command
    edit_parser = subparsers.add_parser(
        'edit', 
        help='VBA-Module bearbeiten mit Live-Sync'
    )
    edit_parser.add_argument(
        '--file', '-f', 
        required=True, 
        help='Visio-Datei (.vsdm)'
    )
    edit_parser.add_argument(
        '--output', '-o', 
        help='Export-Verzeichnis (Standard: aktuelles Verzeichnis)'
    )
    
    # export command
    export_parser = subparsers.add_parser(
        'export', 
        help='VBA-Module exportieren'
    )
    export_parser.add_argument(
        '--file', '-f', 
        required=True, 
        help='Visio-Datei'
    )
    export_parser.add_argument(
        '--output', '-o', 
        help='Export-Verzeichnis'
    )
    
    # import command
    import_parser = subparsers.add_parser(
        'import', 
        help='VBA-Module importieren'
    )
    import_parser.add_argument(
        '--file', '-f', 
        required=True, 
        help='Visio-Datei'
    )
    import_parser.add_argument(
        '--input', '-i', 
        help='Import-Verzeichnis'
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
