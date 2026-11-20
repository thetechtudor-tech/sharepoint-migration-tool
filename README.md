# SharePoint Migration Tool

Copies the files listed in an Excel sheet from an SFTP / FTPS / FTP server or a
local folder into a SharePoint Online document library, and fills in each
file's columns from the same sheet. It runs as the signed-in user through
Microsoft Graph.

## Files

- `sp_migrate_ui.py`: browser UI (served only to the local computer)
- `sp_migrate.py`: the engine and a command-line interface
- `build_exe.bat`: builds a single Windows `.exe` with PyInstaller

## Requirements

- Python 3.9 or later
- `pip install msal requests openpyxl paramiko` (`paramiko` is only needed for SFTP)
- An app registration in Microsoft Entra ID: single tenant, platform
  "Mobile and desktop applications" with redirect URI `http://localhost`,
  public client flows allowed, and the delegated Microsoft Graph permission
  `Sites.ReadWrite.All`. You need its client ID and tenant ID.

## Run the UI

```
python sp_migrate_ui.py
```

1. **Where the files are**: SFTP, FTPS, FTP or a folder. Test the connection.
2. **Where they are going**: the two IDs, the site address and the library.
   Sign in to load the library's columns.
3. **Excel sheet and column matching**: choose the workbook, the worksheet and
   the column holding the file name, then match each Excel column to a
   SharePoint column.
4. **Check and run**: Preview changes nothing; Start runs the migration.

## How it behaves

- The Excel sheet decides what moves. Files in the source with no row in the
  sheet are left alone.
- Files are copied. Deleting from the source after a verified upload is an option.
- Files that already exist in SharePoint are skipped unless "replace" is chosen.
- Results are written to `migration_log.csv`.
- Passwords are held in memory only. Other settings are saved to `config.json`.

## Command line

```
python sp_migrate.py init        # write a starter config.json
python sp_migrate.py columns     # list the library's columns
python sp_migrate.py map         # match Excel headers to columns
python sp_migrate.py run --dry-run
python sp_migrate.py run [--overwrite] [--move]
```

## Build a Windows executable

Run `build_exe.bat` on Windows, or:

```
python -m pip install --upgrade pyinstaller msal requests openpyxl paramiko
python -m PyInstaller --onefile --clean --noconfirm --name SharePointMigration sp_migrate_ui.py
```

The program is written to `dist\SharePointMigration.exe`.

## Status

The UI, Excel reading, column matching and preview are tested with sample
files and a local folder. The SFTP, FTP and SharePoint connections have not
yet been run against real servers.

Person, lookup and managed metadata columns are not supported yet.
