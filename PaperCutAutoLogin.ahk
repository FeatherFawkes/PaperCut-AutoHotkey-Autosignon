#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

SendMode "Event"
SetKeyDelay 75, 75

SetTitleMatchMode 3
CoordMode "Mouse", "Window"

global CredentialHelper := "C:\Users\Edison\PaperCut-AHK\Get-PaperCutCredential.ps1"
global SaveCredentialScript := "C:\Users\Edison\PaperCut-AHK\Save-PaperCutCredential.ps1"
global PowerShellExe := A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"

global LoginWindowTitle := "Login"
global LoginWindowClass := "SunAwtDialog"
global LoginWindowExe := "pc-client.exe"
global BalanceTitlePrefix := "Balance for "

global LastLoginHandle := 0
global LastLoginTick := 0
global FocusAttemptCount := 0
global MaxFocusAttempts := 3

EnsureCredentialExists()
SetTimer WatchPaperCutWindows, 500
A_IconTip := "PaperCut auto login is watching for login and balance windows."

^!c::ShowMouseWindowCoords()

WatchPaperCutWindows() {
    global BalanceTitlePrefix, LoginWindowTitle, LoginWindowClass, LoginWindowExe
    global LastLoginHandle, LastLoginTick, FocusAttemptCount, MaxFocusAttempts

    CloseBalanceWindow()

    hwnd := WinExist(LoginWindowTitle " ahk_class " LoginWindowClass " ahk_exe " LoginWindowExe)
    if !hwnd {
        FocusAttemptCount := 0
        return
    }

    now := A_TickCount
    if (hwnd != LastLoginHandle) {
        FocusAttemptCount := 0
    }

    if (hwnd = LastLoginHandle && now - LastLoginTick < 5000) {
        return
    }

    if !WinActive("ahk_id " hwnd) {
        if (FocusAttemptCount >= MaxFocusAttempts) {
            return
        }
        if !WinActivate("ahk_id " hwnd) {
            FocusAttemptCount += 1
            return
        }
        if !WinWaitActive("ahk_id " hwnd, , 1) {
            FocusAttemptCount += 1
            return
        }
        FocusAttemptCount += 1
        Sleep 250
    } else {
        Sleep 250
    }

    if !WinActive("ahk_id " hwnd) {
        return
    }

    if !FillLoginWindow(hwnd) {
        return
    }

    LastLoginHandle := hwnd
    LastLoginTick := now
    FocusAttemptCount := 0
}

FillLoginWindow(hwnd) {
    creds := GetCredentialLines()
    if (creds.Length < 2) {
        TrayTip "PaperCut", "Could not read the saved PaperCut credential.", 5
        return false
    }

    username := creds[1]
    password := creds[2]

    Send "^a"
    Sleep 80
    SendText username
    Sleep 120

    Send "{Tab}"
    Sleep 120
    Send "^a"
    Sleep 80
    SendText password
    Sleep 150

    Send "{Enter}"
    return true
}

CloseBalanceWindow() {
    global BalanceTitlePrefix

    for hwnd in WinGetList("ahk_exe pc-client.exe") {
        title := WinGetTitle("ahk_id " hwnd)
        if (SubStr(title, 1, StrLen(BalanceTitlePrefix)) = BalanceTitlePrefix) {
            WinClose "ahk_id " hwnd
        }
    }
}

GetCredentialLines() {
    global PowerShellExe, CredentialHelper

    shell := ComObject("WScript.Shell")
    command := '"' PowerShellExe '" -NoProfile -ExecutionPolicy Bypass -File "' CredentialHelper '"'
    exec := shell.Exec(command)
    stdout := exec.StdOut.ReadAll()

    if (stdout = "") {
        return []
    }

    stdout := StrReplace(stdout, "`r`n", "`n")
    stdout := StrReplace(stdout, "`r", "`n")
    stdout := RTrim(stdout, "`n")
    return StrSplit(stdout, "`n")
}

EnsureCredentialExists() {
    global PowerShellExe, SaveCredentialScript

    credentialPath := EnvGet("LOCALAPPDATA") "\PaperCutAHK\papercut-credential.xml"
    if FileExist(credentialPath) {
        return
    }

    MsgBox "No saved PaperCut credential was found. You will be prompted to save it now.", "PaperCut"
    RunWait '"' PowerShellExe '" -NoProfile -ExecutionPolicy Bypass -File "' SaveCredentialScript '"'
}

ShowMouseWindowCoords() {
    MouseGetPos &mouseX, &mouseY, &windowHwnd
    title := WinGetTitle("ahk_id " windowHwnd)
    MsgBox "X: " mouseX "`nY: " mouseY "`nWindow: " title, "PaperCut Coord Helper"
}
