;#################################################################
;##### CfK -- media control
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

;----- ReMap: Increase Volume with Mouse
^+WheelUp::
    Send, {Volume_Up}
return

;----- ReMap: Decrease Volume with Mouse
^+WheelDown::
    Send, {Volume_Down}
return

;----- ReMap: Next Track with Mouse
^+#RButton::
    Send, {Media_Next}
return

;----- ReMap: Prev Track with Mouse 
^+#LButton::
    Send, {Media_Prev}
return

;----- ReMap: Play/Pause with Mouse
^+#MButton::
    Send, {Media_Play_Pause}
return