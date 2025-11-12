## Usage

### Edit Mode (with live sync)

```bash
# Basic mode (VS Code -> Visio only)
visiowings edit --file document.vsdm

# With bidirectional sync (VS Code <-> Visio)
visiowings edit --file document.vsdm --bidirectional

# Force overwrite Document modules (ThisDocument.cls)
visiowings edit --file document.vsdm --force

# Debug mode for troubleshooting
visiowings edit --file document.vsdm --debug

# Automatically delete Visio modules when .bas/.cls/.frm files are deleted locally
visiowings edit --file document.vsdm --sync-delete-modules

# All options combined
visiowings edit --file document.vsdm --force --bidirectional --sync-delete-modules --debug

# Custom output directory
visiowings edit --file document.vsdm --output ./vba_modules
```

### New Option: Module Deletion Sync

- `--sync-delete-modules`: When enabled, modules are automatically removed from Visio when their corresponding .bas/.cls/.frm files are deleted locally in VS Code.
- Default is **off**; activate explicitly if you want this behavior.

---

## Command Line Options

### `edit` command

| Option                  | Description                                                                 |
|------------------------|-----------------------------------------------------------------------------|
| `--file`, `-f`         | Visio file path (`.vsdm`) - **required**                                     |
| `--output`, `-o`       | Export directory (default: current directory)                                |
| `--force`              | Force overwrite Document modules (ThisDocument.cls)                          |
| `--bidirectional`      | Enable bidirectional sync (Visio <-> VS Code)                               |
| `--debug`              | Enable verbose debug logging                                                |
| `--sync-delete-modules`| Automatically delete Visio modules when local .bas/.cls/.frm files are deleted|

---

## Features
- 🧹 **Automatic module deletion**: If `--sync-delete-modules` is enabled, local file deletes remove the corresponding VBA module from Visio to maintain consistency.

## Example

```bash
# Enable automatic module delete
visiowings edit --file MyDiagram.vsdm --sync-delete-modules
```

Rest of README unchanged...
