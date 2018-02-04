#!/usr/bin/env python

# The code in this file is originally from: https://github.com/sid0/iconrefresher/tree/4a048e9252a7e4f95a2eab892311ba25a5207626
# The code was originally [released to the public
# domain](http://creativecommons.org/publicdomain/zero/1.0/) and has been
# relicensed as part of the sci-bots/winreg-helpers project under the terms of
# the BSD 3-clause license.
#
# See https://github.com/sci-bots/winreg-helpers/blob/master/LICENSE.md
import ctypes
from ctypes import wintypes

# http://msdn.microsoft.com/en-us/library/ms644950
SendMessageTimeout = ctypes.windll.user32.SendMessageTimeoutA
SendMessageTimeout.restype = wintypes.LPARAM  # aka LRESULT
SendMessageTimeout.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
                               wintypes.UINT, wintypes.UINT, ctypes.c_void_p]

# http://msdn.microsoft.com/en-us/library/bb762118
SHChangeNotify = ctypes.windll.shell32.SHChangeNotify
SHChangeNotify.restype = None
SHChangeNotify.argtypes = [wintypes.LONG, wintypes.UINT, wintypes.LPCVOID, wintypes.LPCVOID]

HWND_BROADCAST     = 0xFFFF
WM_SETTINGCHANGE   = 0x001A
SMTO_ABORTIFHUNG   = 0x0002
SHCNE_ASSOCCHANGED = 0x08000000


def refresh_icons():
    SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0, SMTO_ABORTIFHUNG, 5000, None)
    SHChangeNotify(SHCNE_ASSOCCHANGED, 0, None, None)


if __name__ == "__main__":
    refresh_icons()
