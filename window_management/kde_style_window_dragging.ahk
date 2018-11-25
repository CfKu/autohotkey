;#################################################################
;##### CfK -- Easy Window Dragging
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr
    #NoEnv
    SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
    DetectHiddenWindows, On
    SetTitleMatchMode, RegEx
    SetWinDelay, 2
    CoordMode, Mouse
RETURN  ; end of auto-execute section

moveMouseCursorToActiveWindow()
{
	WinGetTitle, Title2, A

	; Activate top window 
	WinActivate, %Title2%

	;WinGetPos, xtemp, ytemp,,, A

	; 10 is speed, 16, 16 is a good place to doubleclick and close window. 
	MouseMove, 16, 16, 10
}

moveActiveWindowToMouseCursor()
{
	WinGet, active_id, ID, A
	WinActivate, ahk_id %active_id%
	WinRestore, ahk_id %active_id%  ; This un-maximizes fullscreen things to prevent UI bug. 
	
	; Mouse screen coords = mouse relative + win coords therefore..
	WinGetPos, win_x, win_y, win_width, win_height, ahk_id %active_id%  ; get active windows location
	MouseGetPos, mouse_x, mouse_y   ; get mouse location 
	
	;; Calculate actual position
	move_x := mouse_x - win_width / 2
	move_y := mouse_y - 10
	
	WinMove, ahk_id %active_id%, , %move_x%, %move_y%  ; move window to mouse 
}

;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;+++++ Easy Window Dragging -- KDE style (requires XP/2k/NT) -- by Jonny (modified without double alt and winkey instead of alt)
;http://www.autohotkey.com/docs/scripts/EasyWindowDrag_%28KDE%29.htm

#LButton::
    ; Get the initial mouse position and window id, and
    ; abort if the window is maximized.
    MouseGetPos,KDE_X1,KDE_Y1,KDE_id
    WinGet,KDE_Win,MinMax,ahk_id %KDE_id%
    If KDE_Win
        return
    ; Get the initial window position.
    WinGetPos,KDE_WinX1,KDE_WinY1,,,ahk_id %KDE_id%
    Loop
    {
        GetKeyState,KDE_Button,LButton,P ; Break if button has been released.
        If KDE_Button = U
            break
        MouseGetPos,KDE_X2,KDE_Y2 ; Get the current mouse position.
        KDE_X2 -= KDE_X1 ; Obtain an offset from the initial mouse position.
        KDE_Y2 -= KDE_Y1
        KDE_WinX2 := (KDE_WinX1 + KDE_X2) ; Apply this offset to the window position.
        KDE_WinY2 := (KDE_WinY1 + KDE_Y2)
        WinMove,ahk_id %KDE_id%,,%KDE_WinX2%,%KDE_WinY2% ; Move the window to the new position.
    }
return

#RButton::
    ; Get the initial mouse position and window id, and
    ; abort if the window is maximized.
    MouseGetPos,KDE_X1,KDE_Y1,KDE_id
        WinGet,KDE_Win,MinMax,ahk_id %KDE_id%
    If KDE_Win
        return
    ; Get the initial window position and size.
    WinGetPos,KDE_WinX1,KDE_WinY1,KDE_WinW,KDE_WinH,ahk_id %KDE_id%
    ; Define the window region the mouse is currently in.
    ; The four regions are Up and Left, Up and Right, Down and Left, Down and Right.
    If (KDE_X1 < KDE_WinX1 + KDE_WinW / 2)
    KDE_WinLeft := 1
    Else
    KDE_WinLeft := -1
    If (KDE_Y1 < KDE_WinY1 + KDE_WinH / 2)
    KDE_WinUp := 1
    Else
    KDE_WinUp := -1
    Loop
    {
        GetKeyState,KDE_Button,RButton,P ; Break if button has been released.
        If KDE_Button = U
            break
        MouseGetPos,KDE_X2,KDE_Y2 ; Get the current mouse position.
        ; Get the current window position and size.
        WinGetPos,KDE_WinX1,KDE_WinY1,KDE_WinW,KDE_WinH,ahk_id %KDE_id%
        KDE_X2 -= KDE_X1 ; Obtain an offset from the initial mouse position.
        KDE_Y2 -= KDE_Y1
        ; Then, act according to the defined region.
        WinMove,ahk_id %KDE_id%,, KDE_WinX1 + (KDE_WinLeft+1)/2*KDE_X2  ; X of resized window
                                , KDE_WinY1 +   (KDE_WinUp+1)/2*KDE_Y2  ; Y of resized window
                                , KDE_WinW  -     KDE_WinLeft  *KDE_X2  ; W of resized window
                                , KDE_WinH  -       KDE_WinUp  *KDE_Y2  ; H of resized window
        KDE_X1 := (KDE_X2 + KDE_X1) ; Reset the initial position for the next iteration.
        KDE_Y1 := (KDE_Y2 + KDE_Y1)
    }
return

#MButton::
	moveActiveWindowToMouseCursor()
return