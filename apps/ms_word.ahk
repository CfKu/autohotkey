;#################################################################
;##### CfK -- App -- MS Word
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

#IfWinActive ahk_class OpusApp
    ; Strg+Shift+w = Toogle warning on saving documents with comments
    ; "Vor dem Drucken, Speichern oder Senden einer Datei mit Überarbeitungen"
    ; "oder Kommentaren warnen"
    ^+w::
        msoTrue := -1
        msoFalse := 0
        
        try {	
            word := ComObjActive("Word.Application")
            show_warning := word.Options.WarnBeforeSavingPrintingSendingMarkup
            if (show_warning == msoTrue) {
                word.Options.WarnBeforeSavingPrintingSendingMarkup := msoFalse
                ShowtoolTip("Show warning before saving: NO", 2000)
            } else {
                word.Options.WarnBeforeSavingPrintingSendingMarkup := msoTrue
                ShowtoolTip("Show warning before saving: YES", 2000)							
            }
        }
    return

	; EDIT TEXT MODE ----------------------------------------------------------
    ; Strg+Shift+e = Toggle track changes
    ^+e::
        try {
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
            bTrackChanges := oWordDoc.TrackRevisions
			; .TrackMoves
			; .TrackFormatting
            if (bTrackChanges == -1) { ; -1 = msoTrue
				oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
				ShowtoolTip("Track changes: NO", 2000)
            } else {
				oWordDoc.TrackRevisions := -1  ; -1 = msoTrue
				ShowtoolTip("Track changes: YES", 2000)							
			}
        }
    return
	
    ; Strg+Shift+r = Toggle Markups show/hide
    ^+r::
        try {
            oWordView := ComObjActive("Word.Application").ActiveWindow.View
            oWordView.RevisionsView := 0 ; 0 = wdRevisionsViewFinal
            ;oWordView.MarkupMode := 0 ; 0 = wdBalloonRevisions
            if (oWordView.ShowRevisionsAndComments == -1) { ; -1 = msoTrue
                oWordView.ShowRevisionsAndComments := 0  ; 0 = msoFalse
				ShowtoolTip("Markups visible: NO", 2000)
            } else {
                oWordView.ShowRevisionsAndComments := -1 ; -1 = msoTrue
				ShowtoolTip("Markups visible: YES", 2000)
			}
        }
    return
	
	; HIGHLIGHT TEXT ----------------------------------------------------------
	; https://msdn.microsoft.com/en-us/library/office/ff195343.aspx enum WdColorIndex
	; 0 = wdNoHighlight, 4 = wdBrightGreen, 7 = wdYellow, 6 = wdRed, 5 = wdPink, 

    ; Strg+Shift+y = Highlight/Markierung löschen
    ^+y::
        try {
			; stop tracking changes (if tracked)
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
			bTrackChanges := oWordDoc.TrackRevisions
			oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
			; remove highlight from text
            ComObjActive("Word.Application").Selection.Range.HighlightColorIndex := 0 ; 0 = wdNoHighlight
			; restore tracking changes status
			oWordDoc.TrackRevisions := bTrackChanges
        }
    return
    ; Strg+Shift+x = Gelb markieren
    ^+x::
        try {
			; stop tracking changes (if tracked)
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
			bTrackChanges := oWordDoc.TrackRevisions
			oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
			; highlight text
            ComObjActive("Word.Application").Selection.Range.HighlightColorIndex := 7 ; 7 = wdYellow
			; restore tracking changes status
			oWordDoc.TrackRevisions := bTrackChanges
        }
    return
    ; Strg+Shift+c= Magenta markieren
    ^+c::
        try {
			; stop tracking changes (if tracked)
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
			bTrackChanges := oWordDoc.TrackRevisions
			oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
			; highlight text
            ComObjActive("Word.Application").Selection.Range.HighlightColorIndex := 5 ; 5 = wdPink
			; restore tracking changes status
			oWordDoc.TrackRevisions := bTrackChanges
        }
    return
    ; Strg+Shift+v = Grün markieren
    ^+v::
        try {
			; stop tracking changes (if tracked)
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
			bTrackChanges := oWordDoc.TrackRevisions
			oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
			; highlight text
            ComObjActive("Word.Application").Selection.Range.HighlightColorIndex := 4 ; 4 = wdBrightGreen
			; restore tracking changes status
			oWordDoc.TrackRevisions := bTrackChanges
        }
    return
    ; Strg+Shift+b = Türkis markieren
    ^+b::
        try {
			; stop tracking changes (if tracked)
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
			bTrackChanges := oWordDoc.TrackRevisions
			oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
			; highlight text
            ComObjActive("Word.Application").Selection.Range.HighlightColorIndex := 3 ; 3 = wdTurquoise
			; restore tracking changes status
			oWordDoc.TrackRevisions := bTrackChanges
        }
    return
    ; Strg+Shift+n = Rot markieren
    ^+n::
        try {
			; stop tracking changes (if tracked)
			oWordDoc := ComObjActive("Word.Application").ActiveDocument
			bTrackChanges := oWordDoc.TrackRevisions
			oWordDoc.TrackRevisions := 0  ; 0 = msoFalse
			; highlight text
            ComObjActive("Word.Application").Selection.Range.HighlightColorIndex := 6 ; 6 = wdRed
			; restore tracking changes status
			oWordDoc.TrackRevisions := bTrackChanges
        }
    return
	
	; INSERT / DELETE ---------------------------------------------------------
    ; Strg+m = Paste as text only
    ^m::
        try {
            ComObjActive("Word.Application").Selection.PasteSpecial(ComObjMissing(),ComObjMissing(),0,ComObjMissing(),2) ; 0 = wdInLine; 2 = wdPasteText
        }
    return
	
    ; Strg+b = Paste as emf
    ^b::
        try {
            ComObjActive("Word.Application").Selection.PasteSpecial(ComObjMissing(),ComObjMissing(),0,ComObjMissing(),9) ; 0 = wdInLine; 9 = wdPasteEnhancedMetafile
        }
    return
    
    ; Strg+n = Paste as png
    ^n::
        try {
            ComObjActive("Word.Application").Selection.PasteSpecial(ComObjMissing(),ComObjMissing(),0,ComObjMissing(),14) ; 0 = wdInLine; 14 = wdPastePNG
        }
    return
	
    ; Strg+Space = Insert thin (not breakable) space U+2009 / 8201 (dezimal)
	; Word: 2009{ALT}{C}
    ^Space::
		msoTrue := -1
        try {
            ComObjActive("Word.Application").Selection.InsertSymbol(0x2009, "", msoTrue)
        }
    return
    
    ; Strg+Super++(Numpad) = Zeile in Word-Tabelle unterhalb hinzufügen
    ^#NumpadAdd::
        try {
            ComObjActive("Word.Application").Selection.InsertRowsBelow(1)
        }
    return
    
    ; Strg+Super+-(Numpad) = Aktuelle Zeile in Word-Tabelle löschen
    ^#NumpadSub::
        try {
            ComObjActive("Word.Application").Selection.Rows.Delete()
        }
    return
	
    ; Strg+Shift+l = Toggle German/English auto correction
    ^+l::
        wdGerman := 1031
		wdEnglishUK := 2057
        wdEnglishUS := 1033
        
        try {
			oSelection := ComObjActive("Word.Application").Selection
			iNewLangID := (oSelection.LanguageID = wdGerman) ? wdEnglishUS : wdGerman
			oSelection.LanguageID := iNewLangID
            ShowToolTip((iNewLangID = wdGerman) ? "Deutsch" : "English", 1500)
        }
    return
	
    ; Strg+Shift+Super+l = Toggle spellchecking
    ^+#l::
        msoTrue := -1
        msoFalse := 0
        
        try {
			oSelection := ComObjActive("Word.Application").Selection
			iNewProofing := (oSelection.NoProofing = msoFalse) ? msoTrue : msoFalse
			oSelection.NoProofing := iNewProofing
            ShowToolTip((iNewProofing = msoFalse) ? "Spellchecking: YES" : "Spellchecking: NO", 1500)
        }
    return
#IfWinActive