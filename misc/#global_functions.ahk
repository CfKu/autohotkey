;#################################################################
;##### CfK -- core -- Global functions
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

ShowToolTip(sText, iDuration)
{
    ToolTip, %sText%
    SetTimer, RemoveToolTip, %iDuration%
}

RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
return

BringToFront(process_id)
{
	; ; Process, wait, %process_id%, 4
    ; ; WinActivate, ahk_pid %process_id%
	; ; WinRestore, ahk_pid %process_id%
	; ; WinSet, Top,, ahk_pid %process_id%
	; WinWait, ahk_pid %process_id%, , 10
	; WinActivate, ahk_pid %process_id%
	; WinSet, Top, , ahk_pid %process_id%
	; Loop, 12
	; {
	; 	if WinExist("ahk_pid " . process_id)
	; 	{
	; 		WinActivate, ahk_pid %process_id%
	; 		WinSet, Top, , ahk_pid %process_id%
	; 		; ShowToolTip("Brought to Front", 2000)
	; 		Break
	; 	}
	; 	; ShowToolTip("failed", 100)
	; 	Sleep, 250
	; }
	return
}

hasValue(haystack, needle) {
    if (!isObject(haystack))
        return false
    if (haystack.Length()==0)
        return false
    for k, v in haystack
        if (v == needle)
            return true
    return false
}

RunGetOutput(cmd) {
	shell := ComObjCreate("WScript.Shell")
	exec := shell.Exec("cmd.exe /q /c " cmd)
	if (A_LastError)
		return 0
	return exec.StdOut.ReadAll()
}

inRDP() {
	RegRead, CN, HKCU, Volatile Environment, CLIENTNAME
	RegRead, SN, HKCU, Volatile Environment, SESSIONNAME
	If (SN == "")
	{
		Loop, Reg, HKEY_CURRENT_USER\Volatile Environment, K
		{
			RegRead, CN, HKCU, Volatile Environment\%a_LoopRegName%, CLIENTNAME
			RegRead, SN, HKCU, Volatile Environment\%a_LoopRegName%, SESSIONNAME
			If (SN <> "")
				break
		}
	}
		
	If (SN == "")
		return ""
	Else
		return CN
}

updateTrayIcon() {
	If A_ISSUSPENDED = 1
		Menu, Tray, Icon, .\#ico\CfKs.ico, 1, 1
	Else If A_ISSUSPAUSED = 1
		Menu, Tray, Icon, .\#ico\CfKp.ico, 1, 1
	Else
		Menu, Tray, Icon, .\#ico\CfK.ico, 1, 1
}

; isWinVisible(WinTitle)
; {
	; WinGet, Style, Style, %WinTitle% 
	; Transform, Result, BitAnd, %Style%, 0x10000000 ; 0x10000000 is WS_VISIBLE. 
	; if (Result <> 0) ;Window is Visible
	; {
		; Return 1
	; }
	; Else  ;Window is Hidden
	; {
		; Return 0
	; }
; }

; Retrieves the size of the monitor, the mouse is present
activeMonitorInfo(ByRef monitor_left, ByRef monitor_top, ByRef monitor_width, ByRef monitor_height) {
	CoordMode, Mouse, Screen
	MouseGetPos, mouse_x , mouse_y
	SysGet, monitor_count, MonitorCount
	Loop %monitor_count%
	{
		SysGet, curMon, Monitor, %a_index%
		if (mouse_x >= curMonLeft and mouse_x <= curMonRight and mouse_y >= curMonTop and mouse_y <= curMonBottom)
		{
			monitor_left := curMonLeft
			monitor_top := curMonTop
			monitor_height := curMonBottom - curMonTop
			monitor_width := curMonRight - curMonLeft
			return
		}
	}
}

; Retrieves the workarea on monitor, the mouse is present
activeMonitorWorkArea(ByRef workarea_left, ByRef workarea_top, ByRef workarea_width, ByRef workarea_height) {
	CoordMode, Mouse, Screen
	MouseGetPos, mouse_x , mouse_y
	SysGet, monitor_count, MonitorCount
	Loop %monitor_count%
	{
		SysGet, curMon, Monitor, %a_index%
		if (mouse_x >= curMonLeft and mouse_x <= curMonRight and mouse_y >= curMonTop and mouse_y <= curMonBottom)
		{
			SysGet, workArea, MonitorWorkArea, %a_index%
			workarea_left := workAreaLeft
			workarea_top := workAreaTop
			workarea_height := workAreaBottom - workAreaTop
			workarea_width := workAreaRight - workAreaLeft
			return
		}
	}
}

; Retrives the size and position of windows taskbar
taskbarInfo(ByRef taskbar_left, ByRef taskbar_top, ByRef taskbar_width, ByRef taskbar_height) {
	WinGetPos, taskbar_left, taskbar_top, taskbar_width, taskbar_height, ahk_class Shell_TrayWnd
}

RunWithProxy(ByRef target) {
	; cmd /V /C "set http_proxy=http://sfl-hza-prx-02.schaeffler.com:8080&&C:\CfK\Clementine\clementine.exe"
	target_call := "SET http_proxy=http://sfl-hza-prx-02.schaeffler.com:8080"
	target_call := target_call . " && SET http_proxys=http://sfl-hza-prx-02.schaeffler.com:8080"
	target_call := target_call . " && SET ftp_proxy=http://sfl-hza-prx-02.schaeffler.com:8080"
	target_call := target_call . " && START " . target
    Run, cmd.exe /c %target_call%, , , NewPID
    BringToFront(NewPID)
}

TriggerWebProxyAuthentication() {
	; US public broadcasting GDPR-compliant site, very small
	URL := "https://text.npr.org/robots.txt"
	try {
		ie := ComObjCreate("InternetExplorer.Application")
		ie.Visible := false ; hide IE
		ie.Navigate(URL)
		; give the browser some time to load the url
		sleep 5000
		; quit the instance of the IE
		ie.quit()
		return 1
	} catch {
		return 0
	}
}
