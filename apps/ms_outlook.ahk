;#################################################################
;##### CfK -- App -- MS Outlook
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

#IfWinActive ahk_class rctrl_renwnd32
    ;Paste Links incl. OSX-Link after copied "Long UNC Path" with PathCopyCopy
    ^b::
		StringReplace, lnkosx, clipboard, \\, smb://
		StringReplace, lnkosx, lnkosx, \, /, , All
		outHTML = 🍏: <a href="%lnkosx%">macOS-Link</a> | <span style="font-family:Wingdings"></span>: <a href="%clipboard%">%clipboard%</a>
		WinClip.SetHTML(outHTML)
		WinClip.Paste()
		Send, {Esc 2}
    return
	
    ; QuickSteps #1 over Numpad without SHIFT // AUSNAMEN ZUVOR ALS GELESEN
    ; DONE-DEL
    ^#LButton::
    #Numpad1::
        Send {CTRLDOWN}{SHIFTDOWN}1{SHIFTUP}{CTRLUP}{UP}
    return
    ^#MButton::
    #Numpad2::
        Send {CTRLDOWN}{SHIFTDOWN}2{SHIFTUP}{CTRLUP}{UP}
    return
    ^#RButton::
    #Numpad3::
        Send {CTRLDOWN}{SHIFTDOWN}3{SHIFTUP}{CTRLUP}{UP}
    return
    #Numpad4::
        Send {CTRLDOWN}{SHIFTDOWN}4{SHIFTUP}{CTRLUP}{UP}
    return
    #Numpad5::
        Send {CTRLDOWN}{SHIFTDOWN}5{SHIFTUP}{CTRLUP}{UP}
    return
    #Numpad6::
        Send {CTRLDOWN}{SHIFTDOWN}6{SHIFTUP}{CTRLUP}{UP}
    return

    ; READ / UNREAD
	^WheelUp::^u
	return
	^WheelDown::^q
	return

    ; Super+Numpad0 = Bring delclined appointments back in calender as free 
    #Numpad0::
        olMail := 43
        olAppointment := 26
        olSelection := 74
        olMeetingRequest := 53
        olFree := 0
        olMeetingDeclined := 4
        olMeetingTentative := 2
        olMeetingResponseNegative := 55
        olMeetingResponsePositive := 56
        olMeetingResponseTentative := 57
        ; evtl.1 send decline (prevent from deleting), 2. change free, no reminder, 3. delete manually
        try {
            ; get current selected olMeetingRequest
			selection := ComObjActive("Outlook.Application").ActiveExplorer.Selection
            if selection.Count == 1 {
                selected_item := selection.Item(1)
                appointment_item := ""
                ; proceed if olMeetingRequest                
                if selected_item.Class == olMeetingRequest {
                    appointment_item := selected_item.GetAssociatedAppointment(True)
                ; proceed if selected appointment in calendar
                } else if (selected_item.Class == olAppointment) {
                    appointment_item := selected_item
                }

                if (appointment_item != "") {
                    appointment_item.Subject := "CfK-DECLINED: " appointment_item.Subject
                    appointment_item.BusyStatus := olFree
                    appointment_item.ReminderSet := False
                    appointment_item.Save()
                }
            }
        }
    return

    ; Super+Numpad+ = Add travel time to appointment
    #NumpadAdd::
        olAppointmentItem := 1
        olAppointment := 26

        ; try {
            app := ComObjActive("Outlook.Application")
            ; get current selected olMeetingRequest
			selection := app.ActiveExplorer.Selection
            if selection.Count == 1 {
                selected_item := selection.Item(1)
                ; proceed if olAppointment                
                if selected_item.Class == olAppointment {
                    InputBox, travel_time, Travel time in minutes, Please enter the travel time in minutes., , 200, 150
                    ; if entered something, continue
                    if (!ErrorLevel) {
                        ; cast string to integer
                        travel_time := travel_time + 0
                        ; WORKAROUND (split datetime string), in ahk_v2 DateParse() or DateAdd()
                        ;                  1234567890123456789
                        ; DateTime Format: dd.MM.yyyy hh:mm:ss
                        selected_start_date := selected_item.Start
                        start_dd := SubStr(selected_start_date, 1, 2)
                        start_Mon := SubStr(selected_start_date, 4, 2)
                        start_yyyy := SubStr(selected_start_date, 7, 4)
                        start_hh := SubStr(selected_start_date, 12, 2)
                        start_min := SubStr(selected_start_date, 15, 2)
                        start_ss := SubStr(selected_start_date, 18, 2) 
                        before_start_date_ahk := start_yyyy . start_Mon . start_dd . start_hh . start_min . start_ss
                        before_start_date_ahk += -1 * travel_time, minutes
                        FormatTime, before_start_date, %before_start_date_ahk%, dd.MM.yyyy hh:mm:ss
                       
                        ; BEFORE
                        item_before := app.CreateItem(olAppointmentItem)
                        item_before.Start := before_start_date
                        item_before.Duration := travel_time
                        item_before.Subject := "TRAVEL TIME | " selected_item.Subject
                        item_before.BusyStatus := selected_item.BusyStatus
                        item_before.Categories := selected_item.Categories
                        item_before.Sensitivity := selected_item.Sensitivity
                        item_before.ReminderSet := True
                        item_before.ReminderMinutesBeforeStart := 0
                        item_before.Save()

                        ; AFTER
                        item_after := app.CreateItem(olAppointmentItem)
                        item_after.Start := selected_item.End
                        item_after.Duration := travel_time
                        item_after.Subject := "TRAVEL TIME | " selected_item.Subject
                        item_after.BusyStatus := selected_item.BusyStatus
                        item_after.Categories := selected_item.Categories
                        item_after.Sensitivity := selected_item.Sensitivity
                        item_after.ReminderSet := True
                        item_after.ReminderMinutesBeforeStart := 0
                        item_after.Save()
                    }
                }     
            }
        ; }
    return
	
	; --- IN OPNENED MAIL WINDOW
    ; Strg+Shift+y = Highlight/Markierung löschen
    ^+y::
        wdNoHighlight := 0
        try {
			oCurrentMail := ComObjActive("Outlook.Application").ActiveInspector.currentItem
			oCurrentWordEditor := oCurrentMail.GetInspector.WordEditor.Application
			; highlight text
            oCurrentWordEditor.Selection.Range.HighlightColorIndex := wdNoHighlight
        }
    return
    ; Strg+Shift+x = Gelb markieren
    ^+x::
        wdYellow := 7
        try {
			oCurrentMail := ComObjActive("Outlook.Application").ActiveInspector.currentItem
			oCurrentWordEditor := oCurrentMail.GetInspector.WordEditor.Application
			; highlight text
            oCurrentWordEditor.Selection.Range.HighlightColorIndex := wdYellow
        }
    return
#IfWinActive