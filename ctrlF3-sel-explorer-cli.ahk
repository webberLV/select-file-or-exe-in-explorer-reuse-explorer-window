#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; ================= ENTRY =================
activeHwnd := WinExist("A")
if !activeHwnd
    ExitApp

title := WinGetTitle("ahk_id " activeHwnd)
pid   := WinGetPID("ahk_id " activeHwnd)

; ---- 1) try path from window title ----
path := ""
if RegExMatch(title, 'i)\b([A-Z]:\\[^:*?"<>|\r\n]+?\.[A-Za-z0-9]{1,12})(?=\s+-\s+|$)', &m)
    path := m[1]
else if RegExMatch(title, 'i)\b([A-Z])\\([^:*?"<>|\r\n]+?\.[A-Za-z0-9]{1,12})(?=\s+-\s+|$)', &m2)
    path := m2[1] ':\' m2[2]

; ---- 2) fallback to process exe ----
if (path = "" || !FileExist(path)) {
    hProc := DllCall("Kernel32\OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr")
    if hProc {
        buf := Buffer(65536)
        size := 32768
        if DllCall("Kernel32\QueryFullProcessImageNameW",
            "Ptr", hProc, "UInt", 0, "Ptr", buf.Ptr, "UIntP", &size)
            path := StrGet(buf, size, "UTF-16")
        DllCall("Kernel32\CloseHandle", "Ptr", hProc)
    }
}

if (path != "" && FileExist(path))
    SelectInExistingExplorer(path)

ExitApp


; ================= CORE =================
SelectInExistingExplorer(fullPath) {
    fullPath := Trim(StrReplace(StrReplace(fullPath, "`r"), "`n"))
    SplitPath fullPath, &itemName, &folderPath
    if (folderPath = "")
        return false

    win := FindTopmostExplorerComWindow()
    if !IsObject(win)
        return false

    if !ExplorerNavigateOnly(win, folderPath)
        return false

    if (itemName != "")
        ExplorerSelectItemInView(win, itemName)

    try WinActivate("ahk_id " win.HWND)
    return true
}


; ================= EXPLORER WINDOW SELECTION =================
FindTopmostExplorerComWindow() {
    exHwnd := FindTopmostExplorerHwnd_GlobalZ()
    if !exHwnd
        return 0

    sh := ComObject("Shell.Application")
    for w in sh.Windows {
        try {
            if (w.HWND = exHwnd)
                return w
        }
    }
    return 0
}

FindTopmostExplorerHwnd_GlobalZ() {
    for hwnd in WinGetList() {
        try {
            if !DllCall("IsWindowVisible", "ptr", hwnd, "int")
                continue
            cls := WinGetClass("ahk_id " hwnd)
            if (cls != "CabinetWClass" && cls != "ExploreWClass")
                continue
            if (StrLower(WinGetProcessName("ahk_id " hwnd)) != "explorer.exe")
                continue
            return hwnd
        }
    }
    return 0
}


; ================= NAVIGATION / SELECTION =================
ExplorerNavigateOnly(win, folderPath) {
    try {
        win.Navigate(folderPath)

        start := A_TickCount
        while (A_TickCount - start < 1200) {
            cur := NormalizePathForCompare(UrlFileToPath(win.LocationURL))
            tgt := NormalizePathForCompare(folderPath)
            if (cur != "" && InStr(cur, tgt) = 1)
                return true
            Sleep 40
        }
        return true
    } catch {
        return false
    }
}

ExplorerSelectItemInView(win, itemName) {
    try {
        start := A_TickCount
        while (A_TickCount - start < 1600) {
            doc := win.Document
            folder := doc.Folder
            it := folder.ParseName(itemName)
            if IsObject(it) {
                doc.SelectItem(it, 1|2|4|8)
                return true
            }
            Sleep 60
        }
        return false
    } catch {
        return false
    }
}


; ================= UTILITIES =================
NormalizePathForCompare(p) {
    return RTrim(StrLower(Trim(p)), "\")
}

UrlFileToPath(url) {
    if !InStr(url, "file:///")
        return ""
    p := SubStr(url, 9)
    return UriDecode(StrReplace(p, "/", "\"))
}

UriDecode(s) {
    try {
        return StrGet(BufferFromUriBytes(s), "UTF-8")
    } catch {
        return StrReplace(s, "%20", " ")
    }
}

BufferFromUriBytes(s) {
    bytes := []
    i := 1
    while (i <= StrLen(s)) {
        ch := SubStr(s, i, 1)
        if (ch = "%" && i+2 <= StrLen(s)) {
            bytes.Push(Integer("0x" SubStr(s, i+1, 2)))
            i += 3
        } else {
            bytes.Push(Ord(ch))
            i += 1
        }
    }
    b := Buffer(bytes.Length)
    for j, v in bytes
        NumPut("UChar", v, b, j-1)
    return b
}
