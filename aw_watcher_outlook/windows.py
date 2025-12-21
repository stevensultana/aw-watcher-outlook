import ctypes
import ctypes.wintypes as wintypes
import os
import win32api
import win32com.client
import win32con
import win32gui
import win32process

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
psapi = ctypes.windll.psapi

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010


def get_outlook_activity():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        explorer = outlook.ActiveExplorer()

        if explorer is None:
            return None

        selection = explorer.Selection
        if selection.Count != 1:
            return None

        item = selection.Item(1)

        return {
            "title": item.Subject,
            "folder": explorer.CurrentFolder.Name,
        }

    except Exception:
        return None


def get_app_path(hwnd):
    """Get application path given hwnd."""
    path = None

    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    process = win32api.OpenProcess(
        0x0400, False, pid
    )  # PROCESS_QUERY_INFORMATION = 0x0400

    try:
        path = win32process.GetModuleFileNameEx(process, 0)
    finally:
        win32api.CloseHandle(process)

    return path


def get_app_name(hwnd):
    """Get application filename given hwnd."""
    path = get_app_path(hwnd)

    if path is None:
        return None

    return os.path.basename(path)


def get_outlook_hwnd():
    def callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            try:
                if get_app_name(hwnd).lower() == "outlook.exe":
                    result.append(hwnd)
            except Exception:
                # if window is running as admin, it's not Outlook.
                pass
        return True

    result = []
    win32gui.EnumWindows(callback, result)
    return result[0] if result else None


def is_outlook_on_top(outlook_hwnd):
    # Get the topmost window
    hwnd = win32gui.GetTopWindow(None)

    while hwnd:
        if hwnd == outlook_hwnd:
            return True  # Outlook is above all other windows
        hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)

    return False


def is_outlook_visible():
    # for the future....
    outlook_hwnd = get_outlook_hwnd()
    active_hwnd = win32gui.GetForegroundWindow()

    # If Outlook is active, it's visible
    if active_hwnd == outlook_hwnd:
        print('active is outlook')
        return True

    # Check if active window is maximized
    placement = win32gui.GetWindowPlacement(active_hwnd)
    is_active_maximized = placement[1] == win32con.SW_MAXIMIZE

    # If active window is maximized and Outlook is below it, Outlook is hidden
    if is_active_maximized and not is_outlook_on_top(outlook_hwnd):
        return False

    # Otherwise Outlook is likely visible
    print('default return true')
    return True


def get_active_process_name():
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None

    return get_app_name(hwnd)
