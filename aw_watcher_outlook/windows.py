import ctypes
import os
import pywintypes
import win32api
import win32com.client
import win32process


def get_outlook_activity() -> dict:
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        explorer = outlook.ActiveExplorer()

        if explorer is None:
            return {}

        selection = explorer.Selection
        if selection.Count != 1:
            return {}

        item = selection.Item(1)

        return {
            "title": item.Subject,
            "folder": explorer.CurrentFolder.Name,
        }

    except Exception:
        return {}


def get_app_path(hwnd) -> str:
    """Get application path given hwnd."""
    # The basic function without all the error handling is simple:
    # 1. get process ID: _, pid = win32process.GetWindowThreadProcessId(hwnd)
    # 2. open process to query info: process = win32api.OpenProcess(...)
    # 3. get our requirement - the process filename: path = win32process.GetModuleFileNameEx(process, 0)

    path = ""

    _, pid = win32process.GetWindowThreadProcessId(hwnd)

    try:
        process = win32api.OpenProcess(0x0400, False, pid)  # PROCESS_QUERY_INFORMATION = 0x0400
    except pywintypes.error as e:
        if e.strerror == 'Access is denied.':
            # probably due to admin window - outlook is probably not admin.
            return ""
        else:
            raise e

    try:
        path = win32process.GetModuleFileNameEx(process, 0)
    finally:
        win32api.CloseHandle(process)

    return path


def get_app_name(hwnd) -> str:
    """Get application filename given hwnd."""
    path = get_app_path(hwnd)

    if path == "":
        return ""

    return os.path.basename(path)


def get_active_process_name() -> str:
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    if not hwnd:
        return ""

    return get_app_name(hwnd)
