;#################################################################
;##### CfK -- APP -- Browser (Firefox, Chromium,. ..)
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

#IfWinActive ahk_class MozillaWindowClass
	;----- ReMap on Mouse: Browser_Back
	^WheelLeft::Browser_Back
	return

	;----- ReMap on Mouse: Browser_Forward
	^WheelRight::Browser_Forward
	return
#IfWinActive

#IfWinActive ahk_class IEFrame
	;----- ReMap on Mouse: Browser_Back
	^WheelLeft::Browser_Back
	return

	;----- ReMap on Mouse: Browser_Forward
	^WheelRight::Browser_Forward
	return
#IfWinActive

#IfWinActive ahk_class Chrome_WidgetWin_1
	;----- ReMap on Mouse: Browser_Back
	^WheelLeft::Browser_Back
	return

	;----- ReMap on Mouse: Browser_Forward
	^WheelRight::Browser_Forward
	return
#IfWinActive
