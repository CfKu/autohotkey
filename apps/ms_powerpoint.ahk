;#################################################################
;##### CfK -- App - MS PowerPoint
;# = Super, ^ = Ctrl, ! = Alt, + = Shift, ^>! = AltGr

NodeAdjustSettingsGUI()  {
	global NodeIdsAdjustEdit, AdjustMode1, AdjustMode2, AdjustMode3, AdjustMode4
	Gui, +LastFound
	GuiHWND := WinExist()           ;--get handle to this gui..

	Gui, Add, Text, , Node ids to be adjusted (comma seperated):`n[first node id will be reference node in case of horizontal and vertical alignment]
	Gui, Add, Edit, vNodeIdsAdjustEdit,
	Gui, Add, Text, , Adjust mode:
	Gui, Add, Radio, Checked vAdjustMode1, horizontal alignment |
	Gui, Add, Radio, vAdjustMode2, vertical alignment –––
	Gui, Add, Radio, vAdjustMode3, horizontal distribution |–|–|
	Gui, Add, Radio, vAdjustMode4, vertical distribution E
	Gui, Add, Button, Default, OK
	Gui, Show

	WinWaitClose, ahk_id %GuiHWND%  ;--waiting for gui to close
	return user_choice         		;--returning value
	
	;-------
	ButtonOK:
		user_choice := []
		; ---node_ids_adjust
		GuiControlGet, node_ids_adjust_str, , NodeIdsAdjustEdit
		node_ids_adjust_array := StrSplit(node_ids_adjust_str, ",", " ,")
		node_ids_adjust := []
		for i, node_id in node_ids_adjust_array {
			node_ids_adjust.Insert(0 + node_id)
		}
		user_choice.Insert(node_ids_adjust)	
		; ---adjust_mode
		GuiControlGet, adjust_mode1, , AdjustMode1
		GuiControlGet, adjust_mode2, , AdjustMode2
		GuiControlGet, adjust_mode3, , AdjustMode3
		GuiControlGet, adjust_mode4, , AdjustMode4
		if (adjust_mode1) {
			user_choice.Insert(1)
		} else if (adjust_mode2) {
			user_choice.Insert(2)
		} else if (adjust_mode3) {
			user_choice.Insert(3)
		} else if (adjust_mode4) {
			user_choice.Insert(4)
		} else {
			user_choice.Insert(-1)
		}
		Gui, Destroy
	return
	;-------
	GuiEscape:
	GuiClose:
		user_choice = -1
		Gui, Destroy
	return
}

#IfWinActive ahk_class PPTFrameClass
	global ppt_shape_corner_radius := 1.0
	global ppt_shape_adjustment1 := 0.0
	global ppt_shape_adjustment2 := 0.0
	global ppt_shape_adjustment3 := 0.0
	global ppt_shape_adjustment4 := 0.0
	global ppt_shape_adjustment5 := 0.0
	global ppt_shape_adjustment6 := 0.0
	global ppt_shape_adjustment7 := 0.0
	global ppt_shape_adjustment8 := 0.0
	
	; Shift+NumpadAdd: Zoom + 100%
	+NumpadAdd::
        try {
			zoom := ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom + 100
			if (zoom > 300)
				zoom := 400
            ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := zoom
        }
	return
	
	; Shift+NumpadSub: Zoom - 100%
	+NumpadSub::
        try {
			zoom := ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom - 100
			if (zoom < 100)
				zoom := 10
            ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := zoom
        }	
	return
	
    ; Strg+m = Paste as text only
    ^m::
		ppPasteText := 7
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.View.PasteSpecial(ppPasteText)
        }
    return
    ; Strg+b = Paste as emf
    ^b::
		ppPasteEnhancedMetafile := 2	
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.View.PasteSpecial(ppPasteEnhancedMetafile)
        }
    return
	
    ; Strg+n = Paste as png
    ^n::
		ppPastePNG := 6
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.View.PasteSpecial(ppPastePNG)
        }
    return
    
    ; Strg+Space = Insert thin (not breakable) space U+2009 / 8201 (dezimal)
    ^Space::
		msoTrue := -1
        try {
			text_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.TextRange
            text_range.InsertSymbol(text_range.Font.Name, 0x2009, msoTrue)
        }
    return

    ; Strg+- = Insert optional hyphen U+0AD / 0173 (dezimal)
	^SC035::
		msoTrue := -1
        try {
			text_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.TextRange
			text_range.InsertSymbol(" ", 0x0AD, msoTrue)
        }
    return
	
    ; Strg+NumpadSub = Insert en dash U+2013 / 0150 (dezimal)
    ^NumpadSub::
		msoTrue := -1
        try {
			text_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.TextRange
            text_range.InsertSymbol(text_range.Font.Name, 0x2013, msoTrue)
			; text_range.InsertSymbol(" ", 0x2013, msoTrue)
        }
    return
	
	; Strg+Shift+c: Special Copy
	^+c::
		; ppSelectionNone = 0; ppSelectionSlides = 1; ppSelectionShapes = 2; ppSelectionText = 3
		ppSelectionShapes := 2
	    ; MsoShapeType
		msoAutoShape := 1
		msoFreeform := 5
		msoPicture := 13  ; msoFreeform and msoPicture not working (picture is beschneiden)
		
		try {
			selection := ComObjActive("PowerPoint.Application").ActiveWindow.Selection

			if (selection.Type == ppSelectionShapes) {
				selected_shape := selection.ShapeRange(1)

				; --- Copy shape corner radius		
				if (selected_shape.Type in (msoAutoShape, msoFreeform, msoPicture)) {				
					if (selected_shape.Width < selected_shape.Height) {
						ppt_shape_corner_radius := selected_shape.Width * selected_shape.Adjustments(1)
					} else {
						ppt_shape_corner_radius := selected_shape.Height * selected_shape.Adjustments(1)
					}
				}
				; Store up to 8 shape adjustments
				Loop, % selected_shape.Adjustments.Count {
					ppt_shape_adjustment%A_Index% := selected_shape.Adjustments(A_Index)
				}
			}
			ShowToolTip("Special copy (Strg+Shift+B for paste)", 1500)
		}
		SendInput, ^+c
	return
	
	; Strg+Shift+b = Special Paste/Apply
	^+b::
		; ppSelectionNone = 0; ppSelectionSlides = 1; ppSelectionShapes = 2; ppSelectionText = 3
		ppSelectionShapes := 2
	    ; MsoShapeType
		msoAutoShape := 1
		msoFreeform := 5
		msoPicture := 13
		
		try {
			selection := ComObjActive("PowerPoint.Application").ActiveWindow.Selection
			if (selection.Type == ppSelectionShapes) {
				selected_shape := selection.ShapeRange(1)
				
				; --- Paste shape corner radius
				if (selected_shape.Type in (msoAutoShape, msoFreeform, msoPicture)) {
					if (selected_shape.Width < selected_shape.Height) {
						selected_shape.Adjustments(1) := ppt_shape_corner_radius / selected_shape.Width
					} else {
						selected_shape.Adjustments(1) := ppt_shape_corner_radius / selected_shape.Height
					}
				}
				; Restore up to 8 shape adjustments (except first, reversed)
				Loop, % selected_shape.Adjustments.Count {
					counterVar := 9 - A_Index
					if (GetKeyState("CapsLock", "T") == 0) and (counterVar != 1)
						continue
					selected_shape.Adjustments(counterVar) := ppt_shape_adjustment%counterVar%
				}
			}
		}
	return
	
    ; Strg+Shift+l = Toggle German/English auto correction
    ^+l::
        msoTrue := -1
        msoFalse := 0
        wdGerman := 1031
		wdEnglishUK := 2057
        wdEnglishUS := 1033
        msoAutoShape := 1
        msoGroup := 6
        msoTextBox := 17
        msoPlaceholder := 14
        
        try { 
            oShapeRange := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            iNewLangID := -1
            if (oShapeRange.Type = msoAutoShape) {
                for oShape in oShapeRange {
                    if (oShape.HasTextFrame = msoTrue) {
                        if (iNewLangID = -1)
                            iNewLangID := (oShape.TextFrame.TextRange.LanguageID = wdGerman) ? wdEnglishUS : wdGerman
                        oShape.TextFrame.TextRange.LanguageID := iNewLangID
                    }
                }
            }
            else if (oShapeRange.Type = msoGroup) {
                Loop, % oShapeRange.GroupItems.Count {
                    oShape := oShapeRange.GroupItems.Item(A_Index)
                    if (oShape.HasTextFrame = msoTrue) {
                        if (iNewLangID = -1)
                            iNewLangID := (oShape.TextFrame.TextRange.LanguageID = wdGerman) ? wdEnglishUS : wdGerman
                        oShape.TextFrame.TextRange.LanguageID := iNewLangID 
                    }
                }
            }
            else if (oShapeRange.Type = msoTextBox) or (oShapeRange.Type = msoPlaceholder) {
                if (oShapeRange.HasTextFrame = msoTrue) {
                    if (iNewLangID = -1)
                        iNewLangID := (oShapeRange.TextFrame.TextRange.LanguageID = wdGerman) ? wdEnglishUS : wdGerman
                    oShapeRange.TextFrame.TextRange.LanguageID := iNewLangID
                }
            }
            ShowToolTip((iNewLangID = wdGerman) ? "Deutsch" : "English", 1500)
        }
    return
	
	; Strg+Shift+j: AdJust shape nodes
	^+j::
		; ppSelectionNone = 0; ppSelectionSlides = 1; ppSelectionShapes = 2; ppSelectionText = 3
		ppSelectionShapes := 2
		msoFreeform := 5
		; MsoEditingType: msoEditingAuto = 0; msoEditingCorner = 1; msoEditingSmooth = 2; msoEditingSymmetric = 3;
		; MsoSegmentType: msoSegmentLine = 0; msoSegmentCurve = 1;
        msoTrue := -1
        msoFalse := 0
		msoBringToFront := 0
		msoTextOrientationHorizontal := 1
		ppAlignLeft := 1
		msoAnchorTop := 1
		color_black := 0
		color_white := 16777215
	
		try {
			selection := ComObjActive("PowerPoint.Application").ActiveWindow.Selection
			if (selection.Type == ppSelectionShapes) {
				selected_shape := selection.ShapeRange(1)

				if (selected_shape.Type == msoFreeform) {
					active_slide_index := ComObjActive("PowerPoint.Application").ActiveWindow.View.Slide.SlideIndex
					active_slide := ComObjActive("PowerPoint.Application").ActivePresentation.Slides(active_slide_index)
					shape_nodes := selected_shape.Nodes
					
					; write the number to each node on slide and store it 
					number_boxes := []
					Loop, % shape_nodes.Count {
						node := shape_nodes.Item(A_Index)
						points := node.Points
						number_box := active_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, points[1, 1], points[1, 2], 1, 1)
						number_box.TextFrame.MarginBottom := 0
						number_box.TextFrame.MarginLeft := 0.3
						number_box.TextFrame.MarginRight := 0
						number_box.TextFrame.MarginTop := 0.3
						number_box.TextFrame.TextRange.ParagraphFormat.Alignment := ppAlignLeft
						number_box.TextFrame.VerticalAnchor := msoAnchorTop
						number_box.TextFrame.WordWrap := msoFalse
						number_box.TextFrame.TextRange.Text := A_Index
						number_box.TextFrame.TextRange.Font.Size := 7
						number_box.TextFrame.TextRange.Font.Color.RGB := color_black
						number_box.TextFrame2.TextRange.Font.Glow.Radius := 30
						number_box.TextFrame2.TextRange.Font.Glow.Color.RGB := color_white
						number_box.TextFrame2.TextRange.Font.Glow.Transparency := 0.1
						number_box.ZOrder(msoBringToFront)
						; store text box
						number_boxes.Insert(number_box)
					}
					; ask user for adjust mode and nodes
					user_choice := NodeAdjustSettingsGUI()
					if (user_choice != -1) {
						node_ids_adjust := user_choice[1]
						adjust_mode := user_choice[2]
								
						; check if node_ids to adjust are in shape_nodes.Count
						nodes_ids_in_range := True
						for i, node_id in node_ids_adjust {
							if (node_id > shape_nodes.Count)
									or (node_id <= 0) {
								nodes_ids_in_range := False
								break
							}
						}
						
						; Ok clear to go for the adjustment
						if (nodes_ids_in_range) {
							; ---ADJUST WITH REFERENCE
							if (adjust_mode == 1) or (adjust_mode == 2) {
								; store reference node id (first one)
								node_id_reference := node_ids_adjust.Remove(1)
								reference_point := shape_nodes.Item(node_id_reference).Points
								
								; adjust depending on mode
								if (adjust_mode == 1) {  ; horizontal alignment
									ref_x := reference_point[1, 1]
									for i, node_id in node_ids_adjust {
										adjust_node_point := shape_nodes.Item(node_id).Points
										cur_y := adjust_node_point[1, 2]
										shape_nodes.SetPosition(node_id, ref_x, cur_y)
									}
								} else if (adjust_mode == 2) {  ; vertical alignment
									ref_y := reference_point[1, 2] 
									for i, node_id in node_ids_adjust {
										adjust_node_point := shape_nodes.Item(node_id).Points
										cur_x := adjust_node_point[1, 1]
										shape_nodes.SetPosition(node_id, cur_x, ref_y)
									}										
								}
							; ---ADJUST WITHOUT REFERENCE
							} else if (adjust_mode == 3) or (adjust_mode == 4) {
								if (node_ids_adjust.MaxIndex() >= 3) {
									if (adjust_mode == 3) {  ; horizontal distribution
										adjust_coord_id := 1  ; x-coordinate
										fix_coord_id := 2 ; y-coordinate
									} else if (adjust_mode == 4) {  ; vertical distribution
										adjust_coord_id := 2  ; y-coordinate
										fix_coord_id := 1  ; x-coordinate
									}
								
									node_pos_id := {}							
									; get node id and positions to be sorted
									for i, node_id in node_ids_adjust {
										adjust_node_point := shape_nodes.Item(node_id).Points
										node_pos := adjust_node_point[1, adjust_coord_id]
										node_pos_id[node_pos] := node_id  ; is then directly sorted by key
									}
									; divide nodes in first and last
									; node_pos_id is automatically sorted by key
									node_first := []
									node_last := []
									node_count := 0
									for node_pos, node_id in node_pos_id {
										if (node_count == 0)
											node_first := [node_pos, node_id]
										node_last := [node_pos, node_id]
										node_count += 1
									}
									; compute distribution distance
									node_delta := Abs(node_last[1] - node_first[1]) / (node_count - 1)
									; move nodes according to computed distance
									i := 1
									for node_pos, node_id in node_pos_id {
										if (node_first[2] == node_id)  ; obmit first node
											continue
										if (node_last[2] == node_id)  ; quit on last node
											break
										; move middle node
										adjust_node_point := shape_nodes.Item(node_id).Points
										fix_pos := adjust_node_point[1, fix_coord_id]
										adjust_pos := node_first[1] + node_delta * i
										if (adjust_mode == 3) {  ; horizontal distribution										
											shape_nodes.SetPosition(node_id, adjust_pos, fix_pos)
										} else if (adjust_mode == 4) {  ; vertical distribution
											shape_nodes.SetPosition(node_id, fix_pos, adjust_pos)
										}
										i += 1
									}
								} else {
									MsgBox, % "ERROR: Nodes to be adjustet must be more than two!"
								}
							}
						} else {
							MsgBox, % "ERROR: Node id(s) out of range!"
						}
					}
					
					; delete number boxes
					for i, number_box in number_boxes {
						number_box.Delete
					}
				}
			}
		}
	return
    
    ; Strg+Shift+h = Hide/Show Slide
    ^+h::
        msoTrue := -1
        msoFalse := 0
        
        try {
            slide_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.SlideRange
            new_slide_visibility := (slide_range.SlideShowTransition.Hidden == msoFalse) ? msoTrue : msoFalse
            slide_range.SlideShowTransition.Hidden := new_slide_visibility
        }
    return

    ; Strg+Shift+t = pure Text --> Remove all margins and disable wordwrap
    ^+t::
        msoTrue := -1
        msoFalse := 0
        msoAutoShape := 1
        msoGroup := 6
        msoTextBox := 17
        msoPlaceholder := 14
		ppAutoSizeNone := 0
		ppAlignCenter := 2
		msoAnchorMiddle := 3
        
        try { 
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            if (shape_range.Type = msoAutoShape) {
                for shape in shape_range {
                    if (shape.HasTextFrame = msoTrue) {
						shape.TextFrame.MarginBottom := 0
						shape.TextFrame.MarginLeft := 0
						shape.TextFrame.MarginRight := 0
						shape.TextFrame.MarginTop := 0
						shape.TextFrame.WordWrap := msoFalse
						shape.TextFrame.AutoSize := ppAutoSizeNone
						shape.TextFrame.TextRange.ParagraphFormat.Alignment := ppAlignCenter
						shape.TextFrame.VerticalAnchor := msoAnchorMiddle
                    }
                }
            }
            else if (shape_range.Type = msoGroup) {
                Loop, % shape_range.GroupItems.Count {
                    shape := shape_range.GroupItems.Item(A_Index)
                    if (shape.HasTextFrame = msoTrue) {
						shape.TextFrame.MarginBottom := 0
						shape.TextFrame.MarginLeft := 0
						shape.TextFrame.MarginRight := 0
						shape.TextFrame.MarginTop := 0
						shape.TextFrame.WordWrap := msoFalse
						shape.TextFrame.AutoSize := ppAutoSizeNone
						shape.TextFrame.TextRange.ParagraphFormat.Alignment := ppAlignCenter
						shape.TextFrame.VerticalAnchor := msoAnchorMiddle
                    }
                }
            }
            else if (shape_range.Type = msoTextBox) or (shape_range.Type = msoPlaceholder) {
                if (shape_range.HasTextFrame = msoTrue) {
					shape_range.TextFrame.MarginBottom := 0
					shape_range.TextFrame.MarginLeft := 0
					shape_range.TextFrame.MarginRight := 0
					shape_range.TextFrame.MarginTop := 0
					shape_range.TextFrame.WordWrap := msoFalse
					shape_range.TextFrame.AutoSize := ppAutoSizeNone
					shape_range.TextFrame.TextRange.ParagraphFormat.Alignment := ppAlignCenter
					shape_range.TextFrame.VerticalAnchor := msoAnchorMiddle
                }
            }
        }
    return
	
    ; Strg+Shift+r = Lock/Unlock aspect Ratio of shape
    ^+r::
        msoTrue := -1
        msoFalse := 0
        
        try { 
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			lock_aspect_ratio := (shape_range.LockAspectRatio == msoTrue) ? msoFalse : msoTrue
			shape_range.LockAspectRatio := lock_aspect_ratio
            ShowToolTip((lock_aspect_ratio == msoTrue) ? "Locked" : "Unlocked", 1500)
        }
    return
	
    ; Strg+Shift+z = Start shape editing
    ^+z::
        Send, {AppsKey}
		Sleep, 100
		Send, b
    return
    
    ; Strg+Home = Move selected object to front
    ^Home::
		msoBringToFront := 0
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange.ZOrder(msoBringToFront)
        }
    return

    ; Strg+End = Move selected object to bottom
    ^End::
		msoSendToBack := 1
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange.ZOrder(msoSendToBack)
        }
    return

    ; Strg+BildUp = Move selected object one layer up
    ^PgUp::
		msoBringForward := 2
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange.ZOrder(msoBringForward)
        }
    return

    ; Strg+BildDown = Move selected object one layer down
    ^PgDn::
		msoSendBackward := 3
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange.ZOrder(msoSendBackward)
        }
    return

    ; Strg+g = Group selected objects
    ^g::
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange.Group().Select()
        }
    return

    ; Strg+Shift+g = Ungroup selected objects
    ^+g::
        try {
            ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange.Ungroup().Select()
        }
    return

    ; Align selected objects to the left
    ^Numpad7::
		msoTrue := -1
		msoFalse := 0
		msoAlignLefts := 0
        try {
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            relative_to_slide_edge := (shape_range.Count = 1) ? msoTrue : msoFalse  ; True: relative to slide edge; False: relative to selected elements
            shape_range.Align(msoAlignLefts, relative_to_slide_edge)
        }
    return
	
    ; Align selected objects to the left (Reference: last selected)
    #^Numpad7::
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			if (shape_range.Count > 1) {
				shape_reference := shape_range.Item(shape_range.Count)
				Loop, % shape_range.Count - 1 {
					shape := shape_range.Item(A_Index)
					shape.Left := shape_reference.Left
				}				
			}
        }
    return

    ; Align selected objects horizontal
    ^Numpad8::
		msoTrue := -1
		msoFalse := 0
		msoAlignCenters := 1
        try {
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            relative_to_slide_edge := (shape_range.Count = 1) ? msoTrue : msoFalse  ; True: relative to slide edge; False: relative to selected elements
            shape_range.Align(msoAlignCenters, relative_to_slide_edge)
        }
    return
	
	; Align selected objects horizontal  (Reference: last selected)
    #^Numpad8::
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			if (shape_range.Count > 1) {
				shape_reference := shape_range.Item(shape_range.Count)
				Loop, % shape_range.Count - 1 {
					shape := shape_range.Item(A_Index)
					shape.Left := shape_reference.Left + (shape_reference.Width - shape.Width) / 2
				}				
			}
        }
    return

    ; Align selected objects to the right
    ^Numpad9::
		msoTrue := -1
		msoFalse := 0
		msoAlignRights := 2
        try {
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            relative_to_slide_edge := (shape_range.Count = 1) ? msoTrue : msoFalse  ; True: relative to slide edge; False: relative to selected elements
            shape_range.Align(msoAlignRights, relative_to_slide_edge)
        }
    return
	
    ; Align selected objects to the right (Reference: last selected)
    #^Numpad9::
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			if (shape_range.Count > 1) {
				shape_reference := shape_range.Item(shape_range.Count)
				Loop, % shape_range.Count - 1 {
					shape := shape_range.Item(A_Index)
					shape.Left := shape_reference.Left + shape_reference.Width - shape.Width
				}				
			}
        }
    return

    ; Align selected objects to the top
    ^Numpad4::
		msoTrue := -1
		msoFalse := 0
		msoAlignTops := 3
        try {
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            relative_to_slide_edge := (shape_range.Count = 1) ? msoTrue : msoFalse  ; True: relative to slide edge; False: relative to selected elements
            shape_range.Align(msoAlignTops, relative_to_slide_edge)
        }
    return
	
	; Align selected objects to the top  (Reference: last selected)
    #^Numpad4::
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			if (shape_range.Count > 1) {
				shape_reference := shape_range.Item(shape_range.Count)
				Loop, % shape_range.Count - 1 {
					shape := shape_range.Item(A_Index)
					shape.Top := shape_reference.Top
				}				
			}
        }
    return

    ; Align selected objects centered vertically
    ^Numpad1::
		msoTrue := -1
		msoFalse := 0
		msoAlignMiddles := 4
        try {
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            relative_to_slide_edge := (shape_range.Count = 1) ? msoTrue : msoFalse  ; True: relative to slide edge; False: relative to selected elements
            shape_range.Align(msoAlignMiddles, relative_to_slide_edge)
        }
    return
	
	; Align selected objects centered vertically (Reference: last selected)
    #^Numpad1::
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			if (shape_range.Count > 1) {
				shape_reference := shape_range.Item(shape_range.Count)
				Loop, % shape_range.Count - 1 {
					shape := shape_range.Item(A_Index)
					shape.Top := shape_reference.Top + (shape_reference.Height - shape.Height) / 2
				}				
			}
        }
    return

    ; Align selected objects to the bottom
    ^Numpad0::
		msoTrue := -1
		msoFalse := 0
		msoAlignBottoms := 5
        try {
            shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            relative_to_slide_edge := (shape_range.Count = 1) ? msoTrue : msoFalse  ; True: relative to slide edge; False: relative to selected elements
            shape_range.Align(msoAlignBottoms, relative_to_slide_edge)
        }
    return
	
	; Align selected objects to the bottom  (Reference: last selected)
    #^Numpad0::
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
			if (shape_range.Count > 1) {
				shape_reference := shape_range.Item(shape_range.Count)
				Loop, % shape_range.Count - 1 {
					shape := shape_range.Item(A_Index)
					shape.Top := shape_reference.Top + shape_reference.Height - shape.Height
				}				
			}
        }
    return

    ; Distribute selected elements horizontally to each other
    ^Numpad2::
		msoFalse := 0
		msoDistributeHorizontally := 0		
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            shape_range.Distribute(msoDistributeHorizontally, msoFalse)  ; False: distritbute elements relative
        }
    return

    ; Distribute selected elements vertically to each other
    ^Numpad6::
		msoFalse := 0
		msoDistributeVertically := 1
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            shape_range.Distribute(msoDistributeVertically, msoFalse)  ; False: distritbute elements relative
        }
    return
	
    ; Distribute selected elements horizontally in slide
    ^#Numpad2::
		msoTrue := -1
		msoDistributeHorizontally := 0		
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            shape_range.Distribute(msoDistributeHorizontally, msoTrue)  ; True: distritbute elements relative in slide
        }
    return

    ; Distribute selected elements vertically in slide
    ^#Numpad6::
		msoTrue := -1
		msoDistributeVertically := 1
        try {
			shape_range := ComObjActive("PowerPoint.Application").ActiveWindow.Selection.ShapeRange
            shape_range.Distribute(msoDistributeVertically, msoTrue)  ; True: distritbute elements relative in slide
        }
    return
#IfWinActive