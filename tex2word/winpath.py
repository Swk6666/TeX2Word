"""Safely add or remove an installation directory on HKCU\\Environment\\Path without length limits."""

import ctypes
import os
import sys
import winreg


ENV_KEY = "Environment"
HWND_BROADCAST = 0xFFFF
WM_SETTINGCHANGE = 0x001A
SMTO_ABORTIFHUNG = 0x0002


def _norm(entry: str) -> str:
    return os.path.normcase(os.path.normpath(entry.strip().strip('"')))


def update_user_path(action: str, directory: str) -> None:
    target = os.path.normpath(directory.strip().strip('"'))
    target_key = _norm(target)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, ENV_KEY, 0,
                        winreg.KEY_READ | winreg.KEY_WRITE) as key:
        try:
            current, reg_type = winreg.QueryValueEx(key, "Path")
        except FileNotFoundError:
            current, reg_type = "", winreg.REG_EXPAND_SZ
        entries = [part for part in current.split(";") if part.strip()]
        if action == "add":
            if not any(_norm(part) == target_key for part in entries):
                entries.append(target)
        elif action == "remove":
            entries = [part for part in entries if _norm(part) != target_key]
        else:
            raise ValueError(f"Unsupported action: {action}")
        updated = ";".join(entries)
        if updated != current:
            winreg.SetValueEx(key, "Path", 0, reg_type if reg_type in (winreg.REG_SZ, winreg.REG_EXPAND_SZ)
                              else winreg.REG_EXPAND_SZ, updated)
    result = ctypes.c_size_t()
    ctypes.windll.user32.SendMessageTimeoutW(
        HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment", SMTO_ABORTIFHUNG, 3000, ctypes.byref(result)
    )


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] in ("add", "remove"):
        update_user_path(sys.argv[1], sys.argv[2])
