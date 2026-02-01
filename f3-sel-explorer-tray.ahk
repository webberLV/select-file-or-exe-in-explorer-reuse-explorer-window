#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; ============================================================
; HOTKEY
; - Runs select logic first.
; - If it fails, re-sends a normal F3 to the foreground app.
; ============================================================
$F3:: {
    SelectActiveThingInExplorer()
}


; ============================================================
; TRAY MENU (CUSTOM ONLY, WITH ICONS)
; ============================================================
InitTrayUI()

InitTrayUI() {
    ; Tray icon: folder.ico in script dir if present, else shell32 icon #4
    ico := A_ScriptDir "\folder.ico"
    try {
        if FileExist(ico)
            TraySetIcon(ico)
        else
            TraySetIcon("shell32.dll", 4)
    }

    A_IconTip := "Explorer Select (F3)"

    A_TrayMenu.Delete()

    A_TrayMenu.Add("Open script folder", OpenScriptFolder_ReuseExplorer)
    A_TrayMenu.SetIcon("Open script folder", "shell32.dll", 4)

    A_TrayMenu.Add("Edit this script", EditThisScript_Default)
    A_TrayMenu.SetIcon("Edit this script", "shell32.dll", 70)

    A_TrayMenu.Add("Reload script", ReloadScript)
    A_TrayMenu.SetIcon("Reload script", "shell32.dll", 239)

    A_TrayMenu.Add()

    A_TrayMenu.Add("Exit", ExitScript)
    A_TrayMenu.SetIcon("Exit", "shell32.dll", 28)
}

; ============================================================
; TRAY COMMANDS
; ============================================================
OpenScriptFolder_ReuseExplorer(*) {
    dir := A_ScriptDir

    try {
        win := FindAnyExplorerWindow()
        if IsObject(win) {
            win.Navigate(dir)
            ForceActivateWindowHard(win.HWND)
            return
        }
    }

    Run('explorer.exe "' dir '"')
}

EditThisScript_Default(*) => Edit()
ReloadScript(*) => Reload()
ExitScript(*) => ExitApp()

; ============================================================
; CORE LOGIC
; ============================================================
SelectActiveThingInExplorer() {
    hwnd := WinExist("A")
    if !hwnd
        return false

    pid   := WinGetPID("ahk_id " hwnd)
    proc  := StrLower(WinGetProcessName("ahk_id " hwnd))
    title := WinGetTitle("ahk_id " hwnd)

    path := ""

    if (proc = "code.exe")
        path := NormalizeDrivePrefix(VSCode_CopyActiveFilePath())
    else
        path := NormalizeDrivePrefix(TryGetPathFromTitle(title))

    if (path != "" && FileExist(path)) {
        ExplorerSelectPreferExisting(path)
        return true
    }

    exe := GetProcessPath(pid)
    if (exe != "" && FileExist(exe)) {
        ExplorerSelectPreferExisting(exe)
        return true
    }

    exe2 := GetProcessPath_WMI(pid)
    if (exe2 != "" && FileExist(exe2)) {
        ExplorerSelectPreferExisting(exe2)
        return true
    }

    return false
}

; ============================================================
; EXPLORER HELPERS
; ============================================================
ExplorerSelectPreferExisting(fullPath) {
    fullPath := Trim(StrReplace(StrReplace(fullPath, "`r"), "`n"))
    if ExplorerSelectInExistingWindow(fullPath)
        return
    Run('explorer.exe /select,"' fullPath '"')
}

ExplorerSelectInExistingWindow(fullPath) {
    try {
        win := FindAnyExplorerWindow()
        if !IsObject(win)
            return false

        SplitPath fullPath, &itemName, &folderPath
        if (itemName = "")
            return ExplorerNavigateOnly(win, fullPath)

        if !ExplorerNavigateOnly(win, folderPath)
            return false

        return ExplorerSelectItemInView(win, itemName)
    } catch {
        return false
    }
}

; Pick the topmost (global Z-order) Explorer window at the moment of use.
FindAnyExplorerWindow() {
    hwnd := FindTopmostExplorerHwnd_GlobalZ()
    if !hwnd
        return 0

    sh := ComObject("Shell.Application")
    for w in sh.Windows {
        try {
            if (w.HWND = hwnd)
                return w
        }
    }
    return 0
}

FindTopmostExplorerHwnd_GlobalZ() {
    ; WinGetList() returns top-level windows in Z-order (topmost first)
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

ExplorerNavigateOnly(win, folderPath) {
    try {
        win.Navigate(folderPath)
        ForceActivateWindowHard(win.HWND)

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
                ForceActivateWindowHard(win.HWND)
                return true
            }
            Sleep 60
        }
        return false
    } catch {
        return false
    }
}

; Hard foreground + raise even when another Explorer is in front.
ForceActivateWindowHard(hwnd) {
    if !hwnd
        return false

    try WinRestore("ahk_id " hwnd)
    try WinShow("ahk_id " hwnd)

    try WinActivate("ahk_id " hwnd)
    if WinActive("ahk_id " hwnd)
        return true

    fg  := DllCall("User32\GetForegroundWindow", "ptr")
    meT := DllCall("User32\GetCurrentThreadId", "uint")
    fgT := fg ? DllCall("User32\GetWindowThreadProcessId", "ptr", fg, "uint*", 0, "uint") : 0
    tgT := DllCall("User32\GetWindowThreadProcessId", "ptr", hwnd, "uint*", 0, "uint")

    if (fgT && fgT != meT)
        DllCall("User32\AttachThreadInput", "uint", meT, "uint", fgT, "int", 1)
    if (tgT && tgT != meT)
        DllCall("User32\AttachThreadInput", "uint", meT, "uint", tgT, "int", 1)

    static SWP_NOMOVE := 0x0002
    static SWP_NOSIZE := 0x0001
    static SWP_SHOWWINDOW := 0x0040
    static HWND_TOPMOST := -1
    static HWND_NOTOPMOST := -2

    DllCall("User32\SetWindowPos", "ptr", hwnd, "ptr", HWND_TOPMOST
        , "int", 0, "int", 0, "int", 0, "int", 0
        , "uint", SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW)

    DllCall("User32\SetWindowPos", "ptr", hwnd, "ptr", HWND_NOTOPMOST
        , "int", 0, "int", 0, "int", 0, "int", 0
        , "uint", SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW)

    DllCall("User32\BringWindowToTop", "ptr", hwnd)
    DllCall("User32\SetForegroundWindow", "ptr", hwnd)
    DllCall("User32\SetFocus", "ptr", hwnd)

    if (tgT && tgT != meT)
        DllCall("User32\AttachThreadInput", "uint", meT, "uint", tgT, "int", 0)
    if (fgT && fgT != meT)
        DllCall("User32\AttachThreadInput", "uint", meT, "uint", fgT, "int", 0)

    return WinActive("ahk_id " hwnd)
}

; ============================================================
; UTILITIES
; ============================================================
NormalizeDrivePrefix(p) {
    if (p = "")
        return ""
    if RegExMatch(p, 'i)^([A-Z])\\', &m)
        return m[1] ':\' SubStr(p, 3)
    return p
}

NormalizePathForCompare(p) {
    return RTrim(StrLower(Trim(p)), "\")
}

TryGetPathFromTitle(title) {
    if RegExMatch(title, 'i)\b([A-Z]:\\.+?\.[A-Za-z0-9]{1,12})(?=\s+-\s+|$)', &m)
        return m[1]
    if RegExMatch(title, 'i)\b([A-Z])\\(.+?\.[A-Za-z0-9]{1,12})(?=\s+-\s+|$)', &m2)
        return m2[1] ':\' m2[2]
    return ""
}

VSCode_CopyActiveFilePath() {
    old := ClipboardAll()
    A_Clipboard := ""
    Send "+!c"
    if !ClipWait(0.8) {
        A_Clipboard := old
        return ""
    }
    p := Trim(A_Clipboard)
    A_Clipboard := old
    return p
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

GetProcessPath(pid) {
    static Q := 0x1000
    h := DllCall("Kernel32\OpenProcess", "UInt", Q, "Int", 0, "UInt", pid, "Ptr")
    if !h
        return ""
    try {
        buf := Buffer(65536)
        size := 32768
        ok := DllCall("Kernel32\QueryFullProcessImageNameW", "Ptr", h, "UInt", 0, "Ptr", buf.Ptr, "UIntP", &size)
        return ok ? StrGet(buf, size, "UTF-16") : ""
    } finally {
        DllCall("Kernel32\CloseHandle", "Ptr", h)
    }
}

GetProcessPath_WMI(pid) {
    try {
        wmi := ComObjGet("winmgmts:")
        for p in wmi.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE ProcessId=" pid)
            return p.ExecutablePath
    }
    return ""
}
