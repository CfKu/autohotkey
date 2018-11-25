;#################################################################
;##### CfK -- register user fonts
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr
#NoEnv

; FONT location (linux-like): ~/.fonts
; Source: https://autohotkey.com/board/topic/31031-portable-font/
EnvGet, HOME, HOME
global USER_FONT_LOCATION := HOME "\.fonts"

; Register all fonts in USER_FONT_LOCATION
Loop, %USER_FONT_LOCATION%\*.*
{
	DllCall("AddFontResource", Str, A_LoopFileFullPath)
	; DllCall("RemoveFontResource", Str, A_LoopFileFullPath)
}
; Broadcast the change
PostMessage, 0x1D,,,, ahk_id 0xFFFF