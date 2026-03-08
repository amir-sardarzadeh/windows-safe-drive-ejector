# Windows Safe Drive Ejector

Force-eject external HDD/SSD/USB drives on Windows, with built-in safety guards that block `C:` and any drive on the same physical disk as `C:`.

## Suggested GitHub Repo Name

`windows-safe-drive-ejector`

## Suggested GitHub Description

`Windows tool to safely force-eject external drives, with protection against ejecting C: or partitions on the system disk.`

## Features

- Interactive drive input (`D`, `D:`, `D:\`) or CLI argument mode.
- Normal Explorer eject first, then forced ejection fallback.
- Low-level Windows ejection path using `DeviceIoControl`.
- Extra forced dismount fallback via `mountvol /p`.
- Safety guard blocks:
  - `C:`
  - Any drive on the same physical disk as `C:`
- Admin-ready EXE build (`--uac-admin`).

## Project Files

- `Eject.py` - main Python script.
- `EjectDrive.exe` - built Windows executable.

## Requirements

- Windows 10/11
- Python 3.10+ (for running `.py` directly)
- Administrator privileges recommended for forced ejection paths

## Usage

### EXE (recommended)

1. Run `EjectDrive.exe`.
2. Accept the UAC admin prompt.
3. Enter the drive letter you want to eject (example: `E`).

### Python script

Interactive:

```powershell
python Eject.py
```

Argument mode:

```powershell
python Eject.py E:
```

## Build EXE

```powershell
python -m PyInstaller --noconfirm --onefile --console --name EjectDrive --uac-admin Eject.py
```

## Safety Notes

- This tool refuses to eject `C:` and any partition located on the same physical disk as `C:`.
- Forced ejection can interrupt active writes. Close files/apps using the target drive before ejecting.
