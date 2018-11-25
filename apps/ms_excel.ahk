;#################################################################
;##### CfK -- App -- MS Excel
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

#IfWinActive ahk_class XLMAIN
    ; Strg+b = Paste as formala and number format
    ^b::
        xlPasteFormulasAndNumberFormats := 11
        try {
            ComObjActive("Excel.Application").ActiveSheet.PasteSpecial(xlPasteFormulasAndNumberFormats)
        }
    return

    ; Strg+n = Paste as text only and number format
    ^n::
        xlPasteValuesAndNumberFormats := 12
        try {
            ComObjActive("Excel.Application").ActiveSheet.PasteSpecial(xlPasteValuesAndNumberFormats)
        }
    return

    ; Strg+Super+b = Paste as emf
    ^#b::
        try {
            ComObjActive("Excel.Application").ActiveSheet.PasteSpecial("Bild (Erweiterte Metadatei)")
        }
    return
    
    ; Strg+Super+n = Paste as png
    ^#n::
        try {
            ComObjActive("Excel.Application").ActiveSheet.PasteSpecial("Bild (PNG)")
        }
    return
#IfWinActive