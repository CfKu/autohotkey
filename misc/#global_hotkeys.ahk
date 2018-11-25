;#################################################################
;##### CfK -- Global hotkeys
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

;----- WORKAROUND: Super -> Disable windows menu opening on shortcut
;~LWin Up::
;return
;~RWin Up::
;return

;----- Super+. -> Insert Date Format without -
#.::
    FormatTime, sDate,, yyyy-MM-dd
    SendInput, %sDate%
return

;----- Strg+Super+. -> Insert Date with -
^#.::
    FormatTime, sDate,, yyyyMMdd
    SendInput, %sDate%
return

;----- Strg+Alt+Super+. -> Insert Date with -
^!#.::
    FormatTime, sDate,, yyyy-MM-dd
    WeekOfYear = %A_YDay%
    WeekOfYear /= 7
    WeekOfYear++ ; Convert from 0-base to 1-base 
    SendInput, [DAF] Status Report %sDate% for KW%WeekOfYear%
return

;----- Super+, -> Insert Time  (without :)
#,::
    FormatTime, sTime,, HHmm
    SendInput, %sTime%
return

;----- Super+- -> Insert Date_Time (date without -)
#-::
    FormatTime, sDateTime,, yyyyMMdd--HHmm
    SendInput, %sDateTime%
return

;----- Strg+Super+, -> Insert Time
^#,::
    FormatTime, sTime,, HH:mm
    SendInput, %sTime%
return

;----- Strg+Super+- -> Insert Date_Time (date with -)
^#-::
    FormatTime, sDateTime,, yyyy-MM-dd--HHmm
    SendInput, %sDateTime%
return

;----- Alt+Strg+Super+Shift+h --> Hide current window
!^+#h::
    WinGetTitle, Title, A
    WinHide, %Title%
return

; ;----- Alt+Strg+Super+Shift+c --> Clear cashes
; !^+#c::
; 	; clementine
; 	EnvGet, USERPROFILE, USERPROFILE
; 	Run, rm -R "%USERPROFILE%\.config\Clementine\spotify-cache\Storage"
; 	Run, rm -R "%USERPROFILE%\.config\Clementine\networkcache"
; 	; pyinstaller
; 	Run, rm -R "%USERPROFILE%\AppData\Roaming\pyinstaller"
; 	; .mediathek3
; 	Run, rm -R "%USERPROFILE%\.mediathek3"
; 	; jupyther
; 	Run, rm -R "%USERPROFILE%\AppData\Roaming\jupyter"
; 	; Downloads
; 	Run, rm -R "%USERPROFILE%\Downloads\*.*"
	
; return

; ;----- Alt+Strg+Super+Shift+l --> Link caches in Local/Roaming Folder to C:\
; !^+#l::
; 	cache_local := ["Microsoft\Outlook", "fontconfig"]
; 	cache_roaming := ["Spotify", "texstudio", "Adobe"]

; 	; Local
; 	EnvGet, USERPROFILE, USERPROFILE
; 	for index, app_name in cache_local {
; 		RunWait, junction -d "%USERPROFILE%\AppData\Local\%app_name%"
; 		RunWait, rm -Rf "%USERPROFILE%\AppData\Local\%app_name%"
; 		RunWait, junction "%USERPROFILE%\AppData\Local\%app_name%" "C:\Benutzerdaten\kuestner\cache_Local\%app_name%"
; 	}

; 	; Roaming
; 	EnvGet, USERPROFILE, USERPROFILE
; 	for index, app_name in cache_roaming {
; 		RunWait, junction -d "%USERPROFILE%\AppData\Roaming\%app_name%"
; 		RunWait, rm -Rf "%USERPROFILE%\AppData\Roaming\%app_name%"
; 		RunWait, junction "%USERPROFILE%\AppData\Roaming\%app_name%" "C:\Benutzerdaten\kuestner\cache_Roaming\%app_name%"
; 	}
; return

;----- Super+q -> hide outlook envelope
;#q::
;	try {
;		; Workaround: Get top item in sent folder and mark it unread and read to hide the outlook envelope
;		oFolderSentMail := ComObjActive("Outlook.Application").Session.GetDefaultFolder(5)  ; 5 = Outlook.OlDefaultFolders.olFolderSentMail
;		oLastItem := oFolderSentMail.Items(oFolderSentMail.Items.Count)
;		oLastItem.UnRead := 1  ; True (mark unread)
;		oLastItem.UnRead := 0  ; False (mark read)
;	}
;return

;----- Super+L -> Mute, Lock screen and turn off monitor
#l::
    ; Send {Volume_Mute} ; unmuted bei gemuted ;(
    ; Send {Volume_Down 50}
    Send, {Media_Stop 2}
    Sleep 500
    Run, %A_WinDir%\System32\rundll32.exe user32.dll`, LockWorkStation
    Sleep 1500
    SendMessage 0x112, 0xF170, 2, , Program Manager
return

;----- Strg+Super+Shift+S -> Toggle Screensaver activation
!^+#s::
    ; Check Screensaver Status
    DllCall("SystemParametersInfo", Int,16, UInt,NULL, "UInt *",bIsScreensaverActive, Int,0)
    If (bIsScreensaverActive)
    {
        DllCall("SystemParametersInfo", Int,17, Int,0, UInt,NULL, Int,2)
        MsgBox Screensaver disabled!
    }
    Else
    {
        DllCall("SystemParametersInfo", Int,17, Int,1, UInt,NULL, Int,2)
        MsgBox Screensaver enabled!
    }
return

;------ Super + H --> detect other autohotkeyscripts
#h::
	DetectHiddenWindows, On
	WinGet, wList, List, ahk_class AutoHotkey
	MsgBox % "AutoHoteky scripts running: " wList
return

;------ Strg+Super+Y --> Show pid of current window
^#y::
	MouseGetPos, mouseX, mouseY, mouseid, mousecontrol
	WinGet, processid, PID, ahk_id %mouseid%
	WinGetClass, winclass, ahk_id %mouseid%
	WinGetTitle, wintitle, ahk_id %mouseid%
	WinGetPos, window_x, window_y, window_width, window_height, ahk_id %mouseid%
	out = procces_id: %processid%`nahk_id: %mouseid%`nahk_class (Class): %winclass%`ncontrol (ClassNN): %mousecontrol%`ntitle: %wintitle%`nmouse coord (x, y): %mouseX%, %mouseY%`nwindow dimensions (x, y, w, h): %window_x%, %window_y%, %window_width%, %window_height%
	WinClip.SetText(out)
	ShowToolTip(out, 7000)
return

;------ Strg+Shift+Super+F5 --> Trigger Web Proxy Authentication
^+#F5::
    ShowToolTip("Triggering Web Proxy...", 4500)
    trigger_ok := TriggerWebProxyAuthentication()
    if (trigger_ok == 1) {
        ShowToolTip("Triggered Web Proxy: OK", 3000)
    } else {
        ShowToolTip("Triggered Web Proxy: !!!FAILED!!!", 3000)
    }
return

; ;----- Super+Pause -> Suspend Windows
; #Pause::
; 	;Run rundll32.exe user32.dll`,LockWorkStation
; 	;Sleep 1000
; 	DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
; return

; ;----- Fast Scrolling
; #WheelDown::
;     Send {WheelDown 3}
;     ; Sleep, 75
; return

; #WheelUp::
;     Send {WheelUp 3}
;     ; Sleep, 75
; return
