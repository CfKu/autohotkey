;#################################################################
;##### CfK -- Alternative Tab
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

; ; ----- Switch between Apps same class
; ^+a::Send, {LWinDOWN}1{LWinUP}
; return
; ^+s::Send, {LWinDOWN}2{LWinUP}
; return
; ^+d::Send, {LWinDOWN}3{LWinUP}
; return

; ; Switching between windows of the same app
; !SC056::    ; Next window Cmd+\ (left from Z)
;     WinGetClass, ActiveClass, A
;     WinGet, WinClassCount, Count, ahk_class %ActiveClass%
;     IF WinClassCount = 1
;         Return
;     Else
;         WinSet, Bottom,, A
;     WinActivate, ahk_class %ActiveClass%
; return

; --- Activate next window of same application
#SC056::  ; Super+< (Key left from Y (or Z on EN keyboards))
	WS_VISIBLE := 0x10000000

	WinGetClass, active_class, A
	WinGet, win_class_count, Count, ahk_class %active_class%
	if (win_class_count = 1)
		return
	else
		WinGet, win_list, List, % "ahk_class " active_class
	
	Loop, % win_list
	{
		index := win_list - A_Index + 1
		next_ahk_id := win_list%index%

		; Check if compatible next window
		WinGet, next_win_style, Style, % "ahk_id " next_ahk_id
		WinGet, next_win_state, MinMax, % "ahk_id " next_ahk_id
		if (next_win_style & WS_VISIBLE)
			; and (next_win_state <> -1)  ; check min/max state
		{
			WinActivate, % "ahk_id " next_ahk_id
			break
		}
	}
	; TODO: Windows Explorer and Chrome
	; see https://autohotkey.com/docs/misc/WinTitle.htm#ahk_group
return

; --- Activate last used window
; global tab_last_active_ahk_id := 0x0
global tab_win_list := ""
global tab_win_list_index := 0

cycleLastUsedWindows() {
	WS_VISIBLE := 0x10000000

	WinGetTitle, active_title, A
	Loop, % tab_win_list
	{
		; index := tab_win_list - A_INDEX + 1
		index := A_INDEX + tab_win_list_index + 1
		next_ahk_id := tab_win_list%index%

		; TODO: https://jacksautohotkeyblog.wordpress.com/2017/05/27/check-window-status-with-winget-exstyle-autohotkey-tip/
		; TODO: Use styles to determine if it is a real visible window

		; Check if compatible next window
		WinGet, next_win_style, Style, % "ahk_id " next_ahk_id
		WinGet, next_win_state, MinMax, % "ahk_id " next_ahk_id
		WinGetTitle, next_title, % "ahk_id " next_ahk_id
		; visible and not minimized and exists
		if (next_win_style & WS_VISIBLE)  ; visible
			and (next_win_state <> -1)  ; not minimized
			and (WinExist("ahk_id " next_ahk_id))  ; should exist
			and (active_title <> next_title)
			; and (next_ahk_id <> tab_last_active_ahk_id)
			and (next_title <> "")
			and (next_title <> "Program Manager")
			and (next_title <> "Microsoft Store")
			and (next_title <> "VirtuaWinMainClass")
			and (!inStr(next_title, "Host für die Windows Shell"))
		{
			; MsgBox, % "break: " active_title " >> "next_active_title
			WinActivate, % "ahk_id " next_ahk_id
			;tab_last_active_ahk_id := next_ahk_id
			tab_win_list_index := index
			break	
		}
	}
}

#SC029 Up::  ; Super+^ (Key above tab)
	WinGet, tab_win_list, List
	tab_win_list_index := 0
	cycleLastUsedWindows()
return

#SC029::  ; Super+^ (Key above tab)
	cycleLastUsedWindows()
return

!#SC029::
	; list al visible "windows" >> TODO: Check Other Styles
	WS_VISIBLE := 0x10000000
	delim := "`r`n"
	WinGet, win_list, List
	Loop, % win_list
	{
		current_ahk_id := win_list%A_Index%

		WinGet, win_style, Style, % "ahk_id " current_ahk_id
		WinGet, win_state, MinMax, % "ahk_id " current_ahk_id
		WinGetTitle, win_title, % "ahk_id " current_ahk_id
		If (win_style & WS_VISIBLE)  ; visible
			and (win_state <> -1)  ; not minimized
			and (WinExist("ahk_id " current_ahk_id))  ; should exist
			and (win_title <> "")
			and (win_title <> "Program Manager")
			and (win_title <> "Microsoft Store")
			and (win_title <> "VirtuaWinMainClass")
			and (!inStr(win_title, "Host für die Windows Shell"))
		{
			out .= ( out="" ? "" : delim ) win_title
		}
	}
	MsgBox, % out
return
