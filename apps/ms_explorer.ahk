;#################################################################
;##### CfK -- App -- MS Windows Explorer
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

getCurrentWindowsExplorerPath()
{
	WinGetClass, sClass
	if (sClass = "CabinetWClass") {
		active_hwnd := WinActive("A")
		try {
			open_windows := ComObjCreate("Shell.Application").Windows
			for win in open_windows {
				try {
					if (win.HWND == active_hwnd) {
						for item in win.Document.SelectedItems {
							return item.Path
						}				
					}
				}
			}
		}
	}
	return ""
}

#IfWinActive ahk_class CabinetWClass
	; Duplicate selected object (file or folder) !! Ctrl+d = Shift+Del (normal)
	^d::Send, {CTRLDOWN}c{CTRLUP}{CTRLDOWN}v{CTRLUP}
	return
#IfWinActive