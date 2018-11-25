;#################################################################
;##### CfK -- extended scrolling - mostly horizontal
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

; PowerPoint horizontal scolling
#IfWinActive ahk_class PPTFrameClass
    +WheelDown::  ; RIGHT
        ScrollStep := 5.0
        com_object := ComObjActive("PowerPoint.Application")
        com_object.ActiveWindow.SmallScroll(0,0,ScrollStep,0)
    return

    +WheelUp::  ; LEFT
        ScrollStep := 5.0
        com_object := ComObjActive("PowerPoint.Application")
        com_object.ActiveWindow.SmallScroll(0,0,0,ScrollStep)        
    return
#IfWinActive

; Excel horizontal scolling
#IfWinActive ahk_class XLMAIN
    +WheelDown::  ; RIGHT
        ScrollStep := 3.0
        com_object := ComObjActive("Excel.Application")
        com_object.ActiveWindow.SmallScroll(0,0,ScrollStep,0)
    return

    +WheelUp::  ; LEFT
        ScrollStep := 3.0
        com_object := ComObjActive("Excel.Application")
        com_object.ActiveWindow.SmallScroll(0,0,0,ScrollStep)        
    return
#IfWinActive

; Outlook calendar horizontal scolling
#IfWinActive ahk_class rctrl_renwnd32
    ; WinGetClass, sClass, ahk_id %hWin%
    MouseGetPos,,,, sClassNN
    if (sClassNN == "DayViewWnd1") ; Outlook Calender Viewport
	{
        +WheelDown::  ; RIGHT
            ; Controlsend, , {Shift UP}{Ctrl DOWN}{Right}{Ctrl UP}, ahk_id %hWin%
            SendInput, {Ctrl DOWN}{Right}{Ctrl UP}
        return

        +WheelUp::  ; LEFT
            ; Controlsend, , {Shift UP}{Ctrl DOWN}{Left}{Ctrl UP}, ahk_id %hWin%
            SendInput, {Ctrl DOWN}{Left}{Ctrl UP}       
        return
    }
#IfWinActive