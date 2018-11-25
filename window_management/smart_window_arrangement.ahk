;#################################################################
;##### CfK -- smart window arrangement
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

;----- Alt+Super+Left: Move/Toggle 1/3 oder 2/3 left
!#Left::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := workarea_height
	new_x := workarea_left
	new_y := workarea_top
	if (new_x == current_x and current_width < workarea_width / 2 and new_y == current_y and new_height == current_height)
	{
	   new_width := new_width * 2
	}
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Right: Move/Toggle 1/3 oder 2/3 right
!#Right::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := workarea_height
	new_x := workarea_left + new_width * 2
	new_y := workarea_top
	if (new_x == current_x and current_width < workarea_width / 2 and new_y == current_y and new_height == current_height)
	{
	   new_x := new_x - new_width
	   new_width := new_width * 2
	}
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Up: Move/Toggle 1/3 oder 2/3 up
!#Up::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := workarea_width
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left
	new_y := workarea_top
	
	if (new_x == current_x and new_width == current_width and new_y == current_y and current_height < workarea_height / 2)
	{
	   new_height := new_height * 2
	}
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Down: Move/Toggle 1/3 oder 2/3 down
!#Down::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := workarea_width
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left
	new_y := workarea_top + new_height * 2
	if (new_x == current_x and new_width == current_width and new_y == current_y and current_height < workarea_height / 2)
	{
	   new_y := new_y - new_height
	   new_height := new_height * 2
	}
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+PageUp: Move window 1/2 top
!#PgUp::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := workarea_width
	new_height := Floor(workarea_height / 2)
	new_x := workarea_left
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+PageDown: Move window 1/2 bottom
!#PgDn::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := workarea_width
	new_height := Floor(workarea_height / 2)
	new_x := workarea_left
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Ins: Move window 1/2 left-top
!#Ins::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 2)
	new_height := Floor(workarea_height / 2)
	new_x := workarea_left
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Del: Move window 1/2 left-bottom
!#Del::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 2)
	new_height := Floor(workarea_height / 2)
	new_x := workarea_left
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Home: Move window 1/2 right-top
!#Home::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 2)
	new_height := Floor(workarea_height / 2)
	new_x := workarea_left + new_width
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+End: Move window 1/2 right-bottom
!#End::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 2)
	new_height := Floor(workarea_height / 2)
	new_x := workarea_left + new_width
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad7: Move window 1/3 left-top
!#Numpad7::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad4: Move window 1/3 left-middle
!#Numpad4::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad1: Move window 1/3 left-bottom
!#Numpad1::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left
	new_y := workarea_top + new_height * 2
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad8: Move window 1/3 center-top
!#Numpad8::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left + new_width
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad5: Move window 1/3 center-middle
!#Numpad5::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left + new_width
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad2: Move window 1/3 center-bottom
!#Numpad2::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left + new_width
	new_y := workarea_top + new_height * 2
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad9: Move window 1/3 right-top
!#Numpad9::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left + new_width * 2
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad6: Move window 1/3 right-middle
!#Numpad6::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left + new_width * 2
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad3: Move window 1/3 right-bottom
!#Numpad3::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left + new_width * 2
	new_y := workarea_top + new_height * 2
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+NumpadAdd: Move window to the 1/3 horizontal center
!#NumpadAdd::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := Floor(workarea_width / 3)
	new_height := workarea_height
	new_x := workarea_left + new_width
	new_y := workarea_top
	WinMove, A,, new_x, new_y, new_width, new_height
return

;----- Alt+Super+Numpad0: Move window to the 1/3 vertical center
!#Numpad0::
	WinGetPos, current_x, current_y, current_width, current_height, A
	activeMonitorWorkArea(workarea_left, workarea_top, workarea_width, workarea_height)
	
	new_width := workarea_width
	new_height := Floor(workarea_height / 3)
	new_x := workarea_left
	new_y := workarea_top + new_height
	WinMove, A,, new_x, new_y, new_width, new_height
return