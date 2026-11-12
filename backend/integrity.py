"""Windows process integrity level detection.

Outlook COM is a cross-process desktop COM interface. Windows blocks COM calls
between different integrity levels, and Outlook normally runs at Medium
integrity while sandboxed/tooling processes often run at Low integrity. When the
call is blocked, Outlook may instead launch a duplicate instance and pop a
"Only one version of Outlook can run at a time" dialog, which then blocks the
caller until it times out.

Detecting the integrity level up front lets the CLI fail in milliseconds with an
actionable message instead of hanging for a minute and leaving a modal dialog on
the user's desktop.
"""

from __future__ import annotations

import os

# Windows integrity level RIDs (winnt.h SECURITY_MANDATORY_*_RID).
SECURITY_MANDATORY_LOW_RID = 0x1000
SECURITY_MANDATORY_MEDIUM_RID = 0x2000

_TOKEN_QUERY = 0x0008
_TOKEN_INTEGRITY_LEVEL = 25  # TOKEN_INFORMATION_CLASS::TokenIntegrityLevel


def current_integrity_rid() -> int | None:
    """Return this process token's integrity level RID, or None when unknown."""
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:  # pragma: no cover - defensive
        return None

    try:
        advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        advapi32.OpenProcessToken.argtypes = [
            wintypes.HANDLE,
            wintypes.DWORD,
            ctypes.POINTER(wintypes.HANDLE),
        ]
        advapi32.OpenProcessToken.restype = wintypes.BOOL
        advapi32.GetTokenInformation.argtypes = [
            wintypes.HANDLE,
            ctypes.c_int,
            ctypes.c_void_p,
            wintypes.DWORD,
            ctypes.POINTER(wintypes.DWORD),
        ]
        advapi32.GetTokenInformation.restype = wintypes.BOOL
        advapi32.GetSidSubAuthorityCount.argtypes = [ctypes.c_void_p]
        advapi32.GetSidSubAuthorityCount.restype = ctypes.POINTER(ctypes.c_ubyte)
        advapi32.GetSidSubAuthority.argtypes = [ctypes.c_void_p, wintypes.DWORD]
        advapi32.GetSidSubAuthority.restype = ctypes.POINTER(wintypes.DWORD)

        token = wintypes.HANDLE()
        if not advapi32.OpenProcessToken(
            kernel32.GetCurrentProcess(), _TOKEN_QUERY, ctypes.byref(token)
        ):
            return None
        try:
            size = wintypes.DWORD(0)
            advapi32.GetTokenInformation(
                token, _TOKEN_INTEGRITY_LEVEL, None, 0, ctypes.byref(size)
            )
            if size.value <= 0:
                return None
            buffer = ctypes.create_string_buffer(size.value)
            if not advapi32.GetTokenInformation(
                token, _TOKEN_INTEGRITY_LEVEL, buffer, size, ctypes.byref(size)
            ):
                return None
            # TOKEN_MANDATORY_LABEL -> SID_AND_ATTRIBUTES.Sid
            sid = ctypes.cast(buffer, ctypes.POINTER(ctypes.c_void_p))[0]
            if not sid:
                return None
            count = advapi32.GetSidSubAuthorityCount(sid).contents.value
            if count < 1:
                return None
            return int(advapi32.GetSidSubAuthority(sid, count - 1).contents.value)
        finally:
            kernel32.CloseHandle(token)
    except Exception:  # pragma: no cover - defensive
        return None


def is_low_integrity() -> bool:
    """True when this process runs below Medium integrity and cannot drive Outlook COM."""
    rid = current_integrity_rid()
    return rid is not None and rid < SECURITY_MANDATORY_MEDIUM_RID
