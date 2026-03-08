import os
import re
import sys
import time
import ctypes
import subprocess
from ctypes import wintypes

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)


def normalize_drive(drive: str) -> str:
    drive = drive.strip().replace('"', "").rstrip("\\").upper()
    if len(drive) == 1 and drive.isalpha():
        drive += ":"
    if not re.fullmatch(r"[A-Z]:", drive):
        raise ValueError("Drive must look like E:")
    return drive


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def eject_with_explorer(drive: str) -> bool:
    # Ask Windows Explorer shell to eject the drive
    ps_script = rf"""
$drive = "{drive}"
$shell = New-Object -ComObject Shell.Application
$myComputer = $shell.Namespace(17)
$item = $myComputer.ParseName($drive)
if ($null -eq $item) {{ exit 2 }}
$item.InvokeVerb("Eject")
Start-Sleep -Milliseconds 1500
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        capture_output=True,
        text=True
    )

    # Give Windows a moment
    time.sleep(1.0)

    # If the root is no longer accessible, we assume success
    return not os.path.exists(drive + "\\")


def _device_io_control(handle, code: int, in_buffer=None) -> bool:
    device_io = kernel32.DeviceIoControl
    device_io.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
        wintypes.LPVOID,
    ]
    device_io.restype = wintypes.BOOL

    bytes_returned = wintypes.DWORD(0)
    if in_buffer is None:
        in_ptr = None
        in_size = 0
    else:
        in_ptr = ctypes.byref(in_buffer)
        in_size = ctypes.sizeof(in_buffer)

    return bool(
        device_io(
            handle,
            code,
            in_ptr,
            in_size,
            None,
            0,
            ctypes.byref(bytes_returned),
            None,
        )
    )


def get_physical_disk_number(drive: str) -> tuple[int | None, str]:
    path = r"\\.\{}".format(drive)

    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value
    IOCTL_STORAGE_GET_DEVICE_NUMBER = 0x002D1080

    class STORAGE_DEVICE_NUMBER(ctypes.Structure):
        _fields_ = [
            ("DeviceType", wintypes.DWORD),
            ("DeviceNumber", wintypes.DWORD),
            ("PartitionNumber", wintypes.DWORD),
        ]

    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL

    device_io = kernel32.DeviceIoControl
    device_io.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
        wintypes.LPVOID,
    ]
    device_io.restype = wintypes.BOOL

    handle = create_file(
        path,
        0,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        None,
        OPEN_EXISTING,
        0,
        None,
    )
    if handle == INVALID_HANDLE_VALUE:
        code = ctypes.get_last_error()
        return None, f"CreateFileW({drive}) failed with WinError {code}"

    try:
        info = STORAGE_DEVICE_NUMBER()
        bytes_returned = wintypes.DWORD(0)
        ok = device_io(
            handle,
            IOCTL_STORAGE_GET_DEVICE_NUMBER,
            None,
            0,
            ctypes.byref(info),
            ctypes.sizeof(info),
            ctypes.byref(bytes_returned),
            None,
        )
        if not ok:
            code = ctypes.get_last_error()
            return None, f"IOCTL_STORAGE_GET_DEVICE_NUMBER failed with WinError {code}"
        return int(info.DeviceNumber), ""
    finally:
        close_handle(handle)


def can_eject_drive(drive: str) -> tuple[bool, str]:
    system_drive = "C:"
    if drive == system_drive:
        return False, "Safety block: refusing to eject C:."

    system_disk, system_err = get_physical_disk_number(system_drive)
    if system_disk is None:
        return False, f"Safety block: cannot resolve C: physical disk ({system_err})."

    target_disk, target_err = get_physical_disk_number(drive)
    if target_disk is None:
        return False, f"Safety block: cannot resolve {drive} physical disk ({target_err})."

    if target_disk == system_disk:
        return False, (
            f"Safety block: {drive} is on physical disk {target_disk}, "
            "the same disk as C:."
        )

    return True, ""


def force_eject_with_deviceiocontrol(drive: str) -> tuple[bool, str]:
    path = r"\\.\{}".format(drive)

    GENERIC_READ = 0x80000000
    GENERIC_WRITE = 0x40000000
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    OPEN_EXISTING = 3
    INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

    FSCTL_LOCK_VOLUME = 0x00090018
    FSCTL_DISMOUNT_VOLUME = 0x00090020
    IOCTL_STORAGE_MEDIA_REMOVAL = 0x002D4804
    IOCTL_STORAGE_EJECT_MEDIA = 0x002D4808

    class PREVENT_MEDIA_REMOVAL(ctypes.Structure):
        _fields_ = [("PreventMediaRemoval", wintypes.BOOLEAN)]

    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL

    handle = create_file(
        path,
        GENERIC_READ | GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        None,
        OPEN_EXISTING,
        0,
        None,
    )
    if handle == INVALID_HANDLE_VALUE:
        code = ctypes.get_last_error()
        return False, f"CreateFileW failed with WinError {code}"

    try:
        locked = False
        for _ in range(12):
            if _device_io_control(handle, FSCTL_LOCK_VOLUME):
                locked = True
                break
            time.sleep(0.25)

        if not locked:
            code = ctypes.get_last_error()
            return False, f"FSCTL_LOCK_VOLUME failed with WinError {code}"

        if not _device_io_control(handle, FSCTL_DISMOUNT_VOLUME):
            code = ctypes.get_last_error()
            return False, f"FSCTL_DISMOUNT_VOLUME failed with WinError {code}"

        allow_removal = PREVENT_MEDIA_REMOVAL(False)
        _device_io_control(handle, IOCTL_STORAGE_MEDIA_REMOVAL, allow_removal)

        if not _device_io_control(handle, IOCTL_STORAGE_EJECT_MEDIA):
            code = ctypes.get_last_error()
            return False, f"IOCTL_STORAGE_EJECT_MEDIA failed with WinError {code}"

        time.sleep(0.75)
        return True, "DeviceIoControl ejection path succeeded"
    finally:
        close_handle(handle)


def force_dismount_with_mountvol(drive: str) -> tuple[bool, str]:
    # /p removes the mount point and dismounts the volume immediately.
    result = subprocess.run(
        ["mountvol", drive + "\\", "/p"],
        capture_output=True,
        text=True,
    )
    if result.returncode == 0:
        time.sleep(0.5)
        return True, "mountvol /p dismounted the volume"

    stderr = (result.stderr or "").strip()
    stdout = (result.stdout or "").strip()
    detail = stderr or stdout or f"mountvol exited with code {result.returncode}"
    return False, detail


def force_eject(drive: str) -> tuple[bool, str]:
    ok, detail = force_eject_with_deviceiocontrol(drive)
    if ok:
        return True, detail

    ok2, detail2 = force_dismount_with_mountvol(drive)
    if ok2:
        return True, detail2

    return False, f"{detail}; fallback failed: {detail2}"


def get_drive_from_user() -> str:
    while True:
        raw = input("Enter drive letter to eject (example: E or E:): ").strip()
        try:
            return normalize_drive(raw)
        except ValueError:
            print("Invalid drive. Use a single letter like E or E:.")


def main():
    interactive = len(sys.argv) == 1
    try:
        if interactive:
            drive = get_drive_from_user()
        elif len(sys.argv) == 2:
            drive = normalize_drive(sys.argv[1])
        else:
            print("Usage: Eject.exe E:")
            return
    except ValueError as exc:
        print(f"Invalid drive: {exc}")
        return
    try:
        # Important: do not stay inside the drive you want to eject
        try:
            os.chdir("C:\\")
        except Exception:
            pass

        allowed, reason = can_eject_drive(drive)
        if not allowed:
            print(reason)
            return

        print(f"Trying to eject {drive} ...")

        ok = eject_with_explorer(drive)
        if ok:
            print(f"{drive} was ejected successfully.")
            return

        print("Explorer eject did not complete. Trying forced ejection...")
        ok, detail = force_eject(drive)
        if ok:
            if os.path.exists(drive + "\\"):
                print(f"{drive} was force-dismounted ({detail}).")
                print("You can unplug it now, but active writes may have been interrupted.")
            else:
                print(f"{drive} was force-ejected ({detail}).")
            return

        print(f"Could not force-eject {drive}.")
        print(f"Details: {detail}")
        print()
        print("Try these:")
        print("1. Close all files on the drive.")
        print("2. Close File Explorer windows showing that drive.")
        print("3. Make sure your terminal is not opened inside that drive.")
        print("4. Pause OneDrive / antivirus if they are scanning it.")
        print("5. Run terminal as Administrator and try again.")
        print()
        if not is_admin():
            print("Note: you are not running as Administrator.")
    finally:
        if interactive:
            input("Press Enter to close...")


if __name__ == "__main__":
    main()
