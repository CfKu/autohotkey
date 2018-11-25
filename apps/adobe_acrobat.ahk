;#################################################################
;##### CfK -- App -- Adobe Acrobat (Professional)
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

#IfWinActive ahk_class AcrobatSDIWindow
	^!k::
		MouseGetPos mouse_x, mouse_y
		; make click in adobe independable of monitor
		activeMonitorInfo(monitor_left, monitor_top, monitor_width, monitor_height)
		click_x := monitor_left + monitor_width - 223
		click_y := monitor_top + 193
		MouseClick, left,  click_x, click_y
		Sleep, 500
		Send, ^a	
		MouseMove %mouse_x%, %mouse_y%
	return


	^NumpadAdd::
		MouseGetPos mouse_x, mouse_y
		; make click in adobe independable of monitor
		activeMonitorInfo(monitor_left, monitor_top, monitor_width, monitor_height)
		click_x := monitor_left + monitor_width - 255
		click_y := monitor_top + 228
		MouseClick, left,  click_x, click_y
		Sleep, 500
		MouseMove %mouse_x%, %mouse_y%
	return

	^NumpadSub::
		MouseGetPos mouse_x, mouse_y
		; make click in adobe independable of monitor
		activeMonitorInfo(monitor_left, monitor_top, monitor_width, monitor_height)
		click_x := monitor_left + monitor_width - 182
		click_y := monitor_top + 228
		MouseClick, left,  click_x, click_y
		Sleep, 500
		Send, ^a{DEL}{ESC}
		MouseMove %mouse_x%, %mouse_y%
	return

	^+NumpadSub::
		MouseGetPos mouse_x, mouse_y
		; make click in adobe independable of monitor
		activeMonitorInfo(monitor_left, monitor_top, monitor_width, monitor_height)
		click_x := monitor_left + monitor_width - 224
		click_y := monitor_top + 228
		MouseClick, left,  click_x, click_y
		Sleep, 500
		MouseMove %mouse_x%, %mouse_y%
	return

	^NumpadMult::
		Send, S
		Sleep, 500
		MouseClick, left
	return
	
	^NumpadDiv::	
		MouseGetPos mouse_x, mouse_y
		; make click in adobe independable of monitor
		activeMonitorInfo(monitor_left, monitor_top, monitor_width, monitor_height)
		click_x := monitor_left + monitor_width - 139
		click_y := monitor_top + 228
		MouseClick, left,  click_x, click_y
		Sleep, 500
		Send, ^a{DEL}{ESC}
		MouseMove %mouse_x%, %mouse_y%
	return
#IfWinActive
