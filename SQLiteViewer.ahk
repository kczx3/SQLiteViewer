#Include <Class_SQLiteDB>
#Include <JSON>
#Include <Toolbar>
#Include <Scintilla>
#Include SQLiteViewer_ResultsListView.ahk
#Include SQLiteViewer_HistoryTab.ahk
#Include SQLiteViewer_Snippets.ahk
#Include SQLiteViewer_DBTreeview.ahk
#Include SQLiteViewer_QueryTabs.ahk

global error_icon := LoadPicture("error.ico", , dummyType)
global successful_icon := LoadPicture("successful.ico", , dummyType)
global loading_icon := LoadPicture("loading.ico", , dummyType)
; global function_icon := "89504E470D0A1A0A0000000D4948445200000018000000180806000000E0773DF8000002D649444154785EED964B4C135D14C77BA765DAA2204291E203241082D5D45A12D1F03046A1088941A42ED0C4951B138C1B23F916AE75A10BDDE85A576A3061652446130C4244DEA820AF48041541BEB4A5C07466FC9FE4343193C16ECACE9BFC72EE3DF74EFEE79CFBC8085DD72D9BD9A4440BFE09D8C808211209FCB9404F6606A5A00FAC8267200548C912D80B3A4027180325C0C9598B6408DC02DBC1637012549A95CD04412412C8068D60114C827510F1FB3CD2D9869AFF826702C39489F1FB738DB58B989B693C5D7D91CB293612A8E152F472344A7DED317761C19E6E4988E6A55FFF5F824F668178B44255B59B42883C70196307B09A09FC008FB81F001F7376B8ECA94EC7537C983FFF6DE1FCCBD7DD6366512D2E2D3F242B49A298056C6602BBC008F72F80FD874B0F3423AA438A127BF2E66DDF17AFEF2E04F430EC686579A99BC7DAF8C44C848FFD9678096D26024EB00FE86000ACD9ED72134D8442E1E730DAD4E7D642AFCFF26268A0A5CC626999A339AFEFDEB8C396B39BFA78DF42A4B3D126FB81154C8165B08E88BC3431FFFDE7288C128EACAC17E5B5BF22DFA9EAAA6B88397768E04A455666868F7CD88B61D221CC048EB0ED072A018174724C4ECF2EA11C3DF82EDAD6DED14ABEAF0BC1EBB29C222853EC531DF956A2D136989899800065DC7FC7021A528E9063A73B3B0D6529A6721415E61791C5D855E2B93D74BCAAACC06AB506354DEBEFEC7A4F020A50A95E4681394EAF9C2F9A8CB3DD8B33AED707AA6E609C0532990CE0AA38EAAF686A087C02E3FE831E1FF9782F25C9107D31C8E5E8A7394D2D1653EF6B9A3E9B9AEABCCA6BD7408409E7BAB31FC4D4585B4FEFE089BEC10FB3FC7629408B673002BAC01DBEB5F57C9B1D40E25B990E5C1CB59DFD82AD0CB68274E37B153FA613A08ED36EE60C56E31BC536CA91F318D151E3315B6A2AF7011A6720836D1C751648639FF8DB4396F8F1E30C4C2254198E8A318C13CFB100A7AB181727834DFF6DF90D3E3FFDF68E2E1DA20000000049454E44AE426082"
; FileOpen("function.png", "r").RawRead(function_icon, FileGetSize("function.png"))

main := GuiCreate("+MaximizeBox +Resize +MinSize800x600", "SQLite Viewer")
main.SetFont("s9", "Tahoma")
main.OnEvent("Close", (gui) => mainClose(gui, snippetsTab))
main.OnEvent("Size", "mainResize")

dbViewerTreeview := new SQLiteViewer_DBTreeview(main, 600, 200)
addQueryTabFn := addQueryTab(db, name) => queryTabs.addQueryTab(db, name)
dbViewerTreeview.OnConnect(addQueryTabFn)
dbViewerTreeview.OnDisconnect((db) => queryTabs.removeQueryTab(db))
dbViewerTreeview.OnNewWorksheet(addQueryTabFn)

global queryTabs := new SQLiteViewer_QueryTabs(main, "h200 w700 ys -Wrap Theme Section vQueryTabs")

runQueryBtn := main.addButton("w" 100 - main.marginX " ys+" queryTabs.tabsInterior.y " Section vRunBtn Disabled", "Run")
runQueryBtn.OnEvent("Click", (ctrl) => (tabIndex := queryTabs.getActiveTab(), runQuery(queryTabs.tabs[tabIndex], queryTabs.getQueryEditText())))

clearResultsBtn := main.addButton("wp xs vClearBtn", "Clear Results")
clearResultsBtn.OnEvent("Click", () => resultsLV.clearRows(true))

utilitiesTabs := main.addTab3("ys-" queryTabs.tabsInterior.y " w200 h600 vUtilities", "Snippets")
TC_EX_GetInterior(utilitiesTabs.hwnd, tabx, taby, tabw, tabh)

snippetsTab := new SQLiteViewer_Snippets(main, tabw - 2, tabh - taby, loadSnippets())

resultsTabs := main.addTab3("w800 h" 600 - queryTabs.tabsInterior.h - main.marginY " xm+" dbViewerTreeview.TV.pos.w + main.MarginX " y" queryTabs.tabsInterior.h + main.MarginY * 2 " vResultsTabs", "Results|History")
TC_EX_GetInterior(resultsTabs.hwnd, tabx, taby, tabw, tabh)
resultsTabs.useTab(1)

resultsLV := new SQLiteViewer_ResultsListView(main, tabh - taby, tabw - 2)

resultsTabs.useTab(2)

historyLV := new SQLiteViewer_HistoryTab(main, tabh - taby - (main.MarginY * 2), (tabw - 2) // 2)

resultsTabs.UseTab()

statusBar := main.addStatusBar("vStatusbar")
statusBar.setParts(200, 600, 200)
main.show()

Hotkey("If", () => isQueryEditFocused(main))
Hotkey("^Enter", () => main.control["query"].visible && runQuery(queryTabs.tabs[queryTabs.getActiveTab()], queryTabs.getQueryEditText()))
Hotkey("If")
return

mainClose(gui, snippetsObj) {
    FileDelete("SQLiteViewer_Snippets.json")
    FileAppend(JSON.dump(snippetsObj.snippets, "", 4), "SQLiteViewer_Snippets.json")
    ExitApp()
}

loadSnippets() {
    return FileExist("SQLiteViewer_Snippets.json") ? JSON.load(FileRead("SQLiteViewer_Snippets.json")) : []
}

isQueryEditFocused(gui) {
    return WinActive("ahk_id " gui.hwnd) && (gui.control["query"].hwnd = gui.FocusedCtrl.hwnd)
}

runQuery(db, sql) {
    global resultsLV, statusBar, historyLV
    static SB_SETICON := 0x40F
    
    SplitPath(db._Path, fileName)
    statusBar.SetText(fileName, 1)
    statusBar.SetText(RegExReplace(sql, "\s+", " "), 2)
    SendMessage(SB_SETICON, 2, loading_icon, statusBar.hwnd) ; Otherwise SB.SetIcon() destroys the previous icon
    statusBar.SetText("Running...", 3)
    statusBar.SetText("", 4)
    
    DllCall("QueryPerformanceFrequency", "Int64*", freq)
    DllCall("QueryPerformanceCounter", "Int64*", queryBefore)
    stat := db.Query(sql, RS)
    DllCall("QueryPerformanceCounter", "Int64*", queryAfter)
    queryTime := Round((queryAfter - queryBefore) / freq * 1000, 3)
    
    if (!stat) {
        historyLV.addRow(fileName, StrReplace(sql, "`n", " `n"), "ERROR")
        SendMessage(SB_SETICON, 2, error_icon, statusBar.hwnd) ; Otherwise SB.SetIcon() destroys the previous icon
        statusBar.SetText("ERROR - " db.ErrorMsg, 3)
        if (historyLV.LV.GetCount() > LV_EX_GetRowsPerPage(historyLV.LV.hwnd)) {
            historyLV.LV.ModifyCol(2, "AutoHdr")
        }
        MsgBox("Code: " db.ErrorCode "`n`nMessage:`n" db.ErrorMsg, "Query Error!", 16)
        return
    }
    
    rowCount := resultsLV.setData(RS)

    ; historyLV.addRow(fileName, StrReplace(sql, "`n", " `n"), rowCount)
    historyLV.addRow(fileName, sql, rowCount)
    if (historyLV.LV.GetCount() > LV_EX_GetRowsPerPage(historyLV.LV.hwnd)) {
        historyLV.LV.ModifyCol(2, "AutoHdr")
    }
    
    SendMessage(SB_SETICON, 2, successful_icon, statusBar.hwnd) ; Otherwise SB.SetIcon() destroys the previous icon
    statusBar.setText(rowCount " rows", 3)
    statusBar.SetText(queryTime "ms elapsed", 4)
}

mainResize(gui, minMax, width, height) {
    static prevDimensions := {}
    static prevMinMax := 0
    static buffer := {w: 0, h: 0, count: 0}
    
    ; If the GUI was previously minimized or was just minimized, then just store that state for the next render
    if (prevMinMax = -1 || minMax = -1) {
        prevMinMax := minMax
        return
    }
    ; Store the initial width and height when the GUI is initially shown
    if (!prevDimensions.HasKey("w")) {
        prevDimensions.w := width, prevDimensions.h := height
        return
    }
    
    ; Add the difference in height and width to our buffer
    buffer.w += width - prevDimensions.w
    buffer.h += height - prevDimensions.h
    buffer.count++
    
    ; Only move and redraw if we've buffered 3 pixels of resizing or we are maximizing or restoring from a maximized state
    if (buffer.count = 3 || minMax = 1 || (prevMinMax = 1 && minMax = 0)) {
        halfH := buffer.h // 2
        remainderH := Mod(buffer.h, 2)
        halfW := buffer.w // 2
        remainderW := Mod(buffer.w, 2)
        
        gui.control["DBtreeview"].move("h" gui.control["DBtreeview"].pos.h + buffer.h)
        gui.control["QueryTabs"].move("h" gui.control["QueryTabs"].pos.h + halfH " w" gui.control["QueryTabs"].pos.w + buffer.w)
        gui.control["Query"].move("h" gui.control["Query"].pos.h + halfH " w" gui.control["Query"].pos.w + buffer.w)
        gui.control["ResultsTabs"].move("y" gui.control["ResultsTabs"].pos.y + halfH " h" gui.control["ResultsTabs"].pos.h + halfH + remainderH " w" gui.control["ResultsTabs"].pos.w + buffer.w)
        gui.control["ResultsLV"].move("h" gui.control["ResultsLV"].pos.h + halfH + remainderH " w" gui.control["ResultsLV"].pos.w + buffer.w)
        gui.control["HistoryLV"].move("h" gui.control["HistoryLV"].pos.h + halfH + remainderH " w" gui.control["HistoryLV"].pos.w + halfW + remainderW)
        
        ; Use the newly computed width of HistoryLV to properly move the X position of HistoryEdit
        gui.control["HistoryEdit"].move("x" gui.control["HistoryLV"].pos.w + gui.MarginX * 2 " h" gui.control["HistoryEdit"].pos.h + halfH + remainderH " w" gui.control["HistoryEdit"].pos.w + halfW)
        gui.control["RunBtn"].move("x" gui.control["RunBtn"].pos.x + buffer.w)
        gui.control["ClearBtn"].move("x" gui.control["ClearBtn"].pos.x + buffer.w)
        gui.control["Utilities"].move("x" gui.control["Utilities"].pos.x + buffer.w " h" gui.control["Utilities"].pos.h + buffer.h)
        gui.control["snippetsTB"].move("", true)
        gui.control["SnippetsLV"].move("h" gui.control["SnippetsLV"].pos.h + buffer.h, true)
        
        ; reset our buffers
        buffer.w := buffer.h := buffer.count := 0
    }
    
    ; regardless, we store the currently passed width and height to be able to compute the diff on the next resize event
    prevDimensions.w := width, prevDimensions.h := height
    
    ; store the current minMax state
    prevMinMax := minMax
}

getCallTip(funcName) {
    static tips := JSON.load(FileRead("SQLite_Calltips.json"))
    
    return tips.HasKey(funcName) ? tips[funcName] : false
}

setupSciControl(sci) {
    sci.SetBufferedDraw(0) ; Scintilla docs recommend turning this off for current systems as they perform window buffering
    sci.SetTechnology(1) ; uses Direct2D and DirectWrite APIs for higher quality

    sci.SetLexer(7) ; SQL
    
    ; Autocomplete
    sci.AutoCSetOrder(1) ; have Scintilla perform the sorting for us
    sci.AutoCSetFillups("", "(", 1)
    ; sci.RGBAImageSetWidth(24)
    ; sci.RGBAImageSetHeight(24)
    ; sci.RegisterRGBAImage(1, function_icon, 1)
    
    ; Indentation
    sci.SetTabWidth(4)
    sci.SetUseTabs(false) ; Indent with spaces
    sci.SetTabIndents(1)
    sci.SetBackspaceUnindents(1) ; Backspace will delete spaces that equal a tab
    sci.SetIndentationGuides(3)
    
    sci.StyleSetFont(sci.STYLE_DEFAULT, "Consolas", 1)
    sci.StyleSetSize(sci.STYLE_DEFAULT, 10)
    sci.StyleSetFore(sci.STYLE_DEFAULT, CvtClr(0xF8F8F2))
    sci.StyleSetBack(sci.STYLE_DEFAULT, CvtClr(0x272822))
    sci.StyleClearAll() ; This message sets all styles to have the same attributes as STYLE_DEFAULT.

    ; Active line background color
    sci.SetCaretLineBack(CvtClr(0x3E3D32))
    sci.SetCaretLineVisible(True)
    sci.SetCaretLineVisibleAlways(1)
    sci.SetCaretFore(CvtClr(0xF8F8F0))

    sci.StyleSetFore(sci.STYLE_LINENUMBER, CvtClr(0xF8F8F2)) ; Margin foreground color
    sci.StyleSetBack(sci.STYLE_LINENUMBER, CvtClr(0x272822)) ; Margin background color

    ; Selection
    Sci.SetSelBack(1, CvtClr(0xBEC0BD))
    sci.SetSelAlpha(80)

    ; sci.StyleSetFore(sci.STYLE_BRACELIGHT, CvtClr(0x3399FF))
    ; sci.StyleSetBold(sci.STYLE_BRACELIGHT, True)

    sci.StyleSetFore(sci.SCE_SQL_COMMENT, CvtClr(0x75715E))
    sci.StyleSetFore(sci.SCE_SQL_COMMENTLINE, CvtClr(0x75715E))
    sci.StyleSetFore(sci.SCE_SQL_COMMENTDOC, CvtClr(0x75715E))
    sci.StyleSetFore(sci.SCE_SQL_COMMENTDOCKEYWORD, CvtClr(0x66D9EF))
    sci.StyleSetFore(sci.SCE_SQL_WORD, CvtClr(0xF92672))
    sci.StyleSetFore(sci.SCE_SQL_NUMBER, CvtClr(0xAE81FF))
    sci.StyleSetFore(sci.SCE_SQL_STRING, CvtClr(0xE6DB74))
    sci.StyleSetFore(sci.SCE_SQL_OPERATOR, CvtClr(0xF92672))
    sci.StyleSetFore(sci.SCE_SQL_USER1, CvtClr(0x66D9EF))

    sci.SetKeywords(0, keywords("keywords"), 1)
    sci.SetKeywords(4, keywords("functions"), 1)

    ; line number margin
    PixelWidth := sci.TextWidth(sci.STYLE_LINENUMBER, "9999", 1)
    sci.SetMarginWidthN(0, PixelWidth)
    sci.SetMarginLeft(0, 2) ; Left padding
    
    ; used as a border between line numbers and content
    borderMarginW := 1
    sci.SetMarginTypeN(1, sci.SC_MARGIN_FORE) ; change the second margin to be of type SC_MARGIN_FORE
    sci.SetMarginWidthN(1, borderMarginW) ; set width to 1 pixel

    sci.SetScrollWidth(sci.ctrl.pos.w - PixelWidth - SysGet(11)) ; Also subtract the width of a vertical scrollbar
}

; ======================================================================================================================
; GetInterior     Retrieves the display area of a tab control relative to it's window.
; Return values:  Always True.
; ======================================================================================================================
TC_EX_GetInterior(HTC, ByRef X, ByRef Y, ByRef W, ByRef H) {
   Static TCM_ADJUSTRECT := 0x1328
   X := Y := W := H := 0
   VarSetCapacity(RECT, 16, 0)
   DllCall("User32.dll\GetClientRect", "Ptr", HTC, "Ptr", &RECT)
   SendMessage(TCM_ADJUSTRECT, 0, &RECT, , "ahk_id " . HTC)
   X := NumGet(RECT, 0, "Int")
   Y := NumGet(RECT, 4, "Int")
   W := NumGet(RECT, 8, "Int")
   H := NumGet(RECT, 12, "Int")
   Return True
}

; ======================================================================================================================
; Add             Adds a new tab at the end of the tabs control.
; Return values:  Returns the 1-based index of the new tab if successful, or 0 otherwise.
; ======================================================================================================================
TC_EX_Add(HTC, TabText, TabIndex := -1, IconIndex := 0) {
    Static TCIF_TEXT := 0x0001
    Static TCIF_IMAGE := 0x0002
    Static TCM_INSERTITEM := A_IsUnicode ? 0x133E : 0x1307 ; TCM_INSERTITEMW : TCM_INSERTITEMA
    Static OffImg := (3 * 4) + (A_PtrSize - 4) + A_PtrSize + 4
    Static OffTxP := (3 * 4) + (A_PtrSize - 4)
    TC_EX_CreateTCITEM(TCITEM)
    Flags := TCIF_TEXT
    If (ItemIcon > 0)
        Flags |= TCIF_IMAGE
    NumPut(Flags, TCITEM, 0, "UInt")
    NumPut(&TabText, TCITEM, OffTxP, "Ptr")
    If (ItemIcon > 0)
        NumPut(IconIndex - 1, TCITEM, OffImg, "Int")
    return SendMessage(TCM_INSERTITEM, TabIndex = -1 ? TC_EX_GetCount(HTC) : TabIndex, &TCITEM, , "ahk_id " . HTC) + 1
}

; ======================================================================================================================
; GetCount        Retrieves the number of tabs in a tab control.
; Return values:  Returns the number of tabs if successful, or zero otherwise.
; ======================================================================================================================
TC_EX_GetCount(HTC) {
    Static TCM_GETITEMCOUNT := 0x1304
    return SendMessage(TCM_GETITEMCOUNT, 0, 0, , "ahk_id " . HTC)
}

; ======================================================================================================================
; CreateTCITEM    >>> For internal use! <<< Creates and initializes a TCITEM structure.
; ======================================================================================================================
TC_EX_CreateTCITEM(ByRef TCITEM) {
   Static Size := (5 * 4) + (2 * A_PtrSize) + (A_PtrSize - 4)
   VarSetCapacity(TCITEM, Size, 0)
}

; ======================================================================================================================
; LV_EX_GetRowsPerPage - Calculates the number of items that can fit vertically in the visible area of a list-view
;                        control when in list or report view. Only fully visible items are counted.
; ======================================================================================================================
LV_EX_GetRowsPerPage(HLV) {
   ; LVM_GETCOUNTPERPAGE = 0x1028 -> http://msdn.microsoft.com/en-us/library/bb774917(v=vs.85).aspx
   return SendMessage(0x1028, 0, 0, , "ahk_id " . HLV)
}

; ======================================================================================================================
; LV_EX_GetHeader - Retrieves the handle of the header control used by the list-view control.
; ======================================================================================================================
LV_EX_GetHeader(HLV) {
   ; LVM_GETHEADER = 0x101F -> http://msdn.microsoft.com/en-us/library/bb774937(v=vs.85).aspx
   return SendMessage(0x101F, 0, 0, , "ahk_id " . HLV)
}

; ======================================================================================================================
; LV_EX_SetTileViewLines - Sets the maximum number of additional text lines in each tile, not counting the title.
; ======================================================================================================================
LV_EX_SetTileViewLines(HLV, Lines, tileX := "", tileY := "") {
	; Lines : Maximum number of text lines in each item label, not counting the title.
	; LVM_GETTILEVIEWINFO = 0x10A3 -> http://msdn.microsoft.com/en-us/library/bb761083(v=vs.85).aspx
	; LVM_SETTILEVIEWINFO = 0x10A2 -> http://msdn.microsoft.com/en-us/library/bb761212(v=vs.85).aspx
	; One line is added internally because the item might be wrapped to two lines!
	Static SizeLVTVI := 40
	Static offSize := 12
	Static OffLines := 20
	Static LVTVIM_TILESIZE := 0x1
	Static LVTVIM_COLUMNS := 0x2
	Static LVTVIF_AUTOSIZE := 0x0, LVTVIF_FIXEDWIDTH := 0x1, LVTVIF_FIXEDHEIGHT := 0x2, LVTVIF_FIXEDSIZE := 0x3
	Mask := LVTVIM_COLUMNS | (tileX || tileY ? LVTVIM_TILESIZE : 0)
	If (tileX && tileY)
		flag := LVTVIF_FIXEDSIZE
	Else If (tileX && !tileY)
		flag := LVTVIF_FIXEDWIDTH
	Else If (!tileX && tileY)
		flag := LVTVIF_FIXEDHEIGHT
	Else
		flag := LVTVIF_AUTOSIZE
	; If (Lines > 0)
	; Lines++
	VarSetCapacity(LVTVI, SizeLVTVI, 0)     ; LVTILEVIEWINFO
	NumPut(SizeLVTVI, LVTVI, 0, "UInt")     ; cbSize
	NumPut(Mask, LVTVI, 4, "UInt")    ; dwMask = LVTVIM_TILESIZE | LVTVIM_COLUMNS
	NumPut(flag, LVTVI, 8, "UInt")       ; dwMask
	if (tileX)
		NumPut(tileX, LVTVI, 12, "Int")       ; sizeTile.cx
	if (tileY)
		NumPut(tileY, LVTVI, 16, "Int")       ; sizeTile.cx
	NumPut(Lines, LVTVI, OffLines, "Int") ; c_lines: max lines below first line
	return SendMessage(0x10A2, 0, &LVTVI, , "ahk_id " . HLV) ; LVM_SETTILEVIEWINFO
}

; ======================================================================================================================
; LV_EX_SubItemHitTest - Gets the column (subitem) at the passed coordinates or the position of the mouse cursor.
; ======================================================================================================================
LV_EX_SubItemHitTest(HLV, X := -1, Y := -1) {
   ; LVM_SUBITEMHITTEST = 0x1039 -> http://msdn.microsoft.com/en-us/library/bb761229(v=vs.85).aspx
   VarSetCapacity(LVHTI, 24, 0) ; LVHITTESTINFO
   If (X = -1) || (Y = -1) {
      DllCall("User32.dll\GetCursorPos", "Ptr", &LVHTI)
      DllCall("User32.dll\ScreenToClient", "Ptr", HLV, "Ptr", &LVHTI)
   }
   Else {
      NumPut(X, LVHTI, 0, "Int")
      NumPut(Y, LVHTI, 4, "Int")
   }
   return SendMessage(0x1039, 0, &LVHTI, , "ahk_id " . HLV) > 0x7FFFFFFF ? 0 : NumGet(LVHTI, 16, "Int") + 1
}

; ======================================================================================================================
; LV_EX_EnableGroupView - Enables or disables whether the items in a list-view control display as a group.
; ======================================================================================================================
LV_EX_EnableGroupView(HLV, Enable := True) {
   ; LVM_ENABLEGROUPVIEW = 0x109D -> msdn.microsoft.com/en-us/library/bb774900(v=vs.85).aspx
   return SendMessage(0x109D, !!Enable, 0, , "ahk_id " . HLV) >> 31 ? 0 : 1
}

; ======================================================================================================================
; LV_EX_GetGroup - Gets the ID of the group the list-view item belongs to.
; ======================================================================================================================
LV_EX_GetGroup(HLV, Row) {
   ; LVM_GETITEMA = 0x1005 -> http://msdn.microsoft.com/en-us/library/bb774953(v=vs.85).aspx
   Static OffGroupID := 28 + (A_PtrSize * 3)
   LV_EX_LVITEM(LVITEM, 0x00000100, Row) ; LVIF_GROUPID
   SendMessage(0x1005, 0, &LVITEM, , "ahk_id " . HLV)
   Return NumGet(LVITEM, OffGroupID, "UPtr")
}

; ======================================================================================================================
; LV_EX_GetGroupHeader - Gets the header text of a group by group ID
; ======================================================================================================================
LV_EX_GetGroupHeader(HLV, GroupID, MaxChars := 1024) {
   ; LVM_GETGROUPINFO = 0x1095
   Static SizeOfLVGROUP := (4 * 6) + (A_PtrSize * 4)
   Static LVGF_HEADER := 0x00000001
   Static OffHeader := 8
   Static OffHeaderMax := 8 + A_PtrSize
   VarSetCapacity(HeaderText, MaxChars * 2, 0)
   VarSetCapacity(LVGROUP, SizeOfLVGROUP, 0)
   NumPut(SizeOfLVGROUP, LVGROUP, 0, "UInt")
   NumPut(LVGF_HEADER, LVGROUP, 4, "UInt")
   NumPut(&HeaderText, LVGROUP, OffHeader, "Ptr")
   NumPut(MaxChars, LVGROUP, OffHeaderMax, "Int")
   SendMessage(0x1095, GroupID, &LVGROUP, , "ahk_id " . HLV)
   Return StrGet(&HeaderText, MaxChars, "UTF-16")
}

; ======================================================================================================================
; LV_EX_GroupInsert - Inserts a group into a list-view control.
; ======================================================================================================================
LV_EX_GroupInsert(HLV, GroupID, Header, Align := "", Index := -1, Subtitle := "") {
    ; LVM_INSERTGROUP = 0x1091 -> msdn.microsoft.com/en-us/library/bb761103(v=vs.85).aspx
    Static Alignment := {1: 1, 2: 2, 4: 4, C: 2, L: 1, R: 4}
    Static SizeOfLVGROUP := (4 * 6) + (A_PtrSize * 5)
    Static OffHeader := 8
    Static OffGroupID := OffHeader + (A_PtrSize * 3) + 4
    Static OffAlign := OffGroupID + 12
    Static OffSubtitle := OffAlign + 4
    Static LVGF_SUBTITLE := 0x00000100
    Static LVGF := 0x11 ; LVGF_GROUPID | LVGF_HEADER | LVGF_STATE
    Static LVGF_ALIGN := 0x00000008
    Align := (A := Alignment[SubStr(Align, 1, 1)]) ? A : 0
    Mask := LVGF | (Align ? LVGF_ALIGN : 0) | (Subtitle ? LVGF_SUBTITLE : 0)
    PHeader := A_IsUnicode ? &Header : LV_EX_PWSTR(Header, WHeader)
    PSubtitle := A_IsUnicode ? &Subtitle : LV_EX_PWSTR(Subtitle, WSubtitle)
    VarSetCapacity(LVGROUP, SizeOfLVGROUP, 0)
    NumPut(SizeOfLVGROUP, LVGROUP, 0, "UInt")
    NumPut(Mask, LVGROUP, 4, "UInt")
    NumPut(PHeader, LVGROUP, OffHeader, "Ptr")
    NumPut(GroupID, LVGROUP, OffGroupID, "Int")
    NumPut(Align, LVGROUP, OffAlign, "UInt")
    NumPut(PSubtitle, LVGROUP, OffSubtitle, "Ptr")
    return SendMessage(0x1091, Index, &LVGROUP, , "ahk_id " . HLV)
}

; ======================================================================================================================
; LV_EX_GroupSetState - Set group state (requires Win Vista+ for most states).
; ======================================================================================================================
LV_EX_GroupSetState(HLV, GroupID, States*) {
   ; LVM_SETGROUPINFO = 0x1093 -> msdn.microsoft.com/en-us/library/bb761167(v=vs.85).aspx
   Static OS := DllCall("GetVersion", "UChar")
   Static LVGS5 := {Collapsed: 0x01, Hidden: 0x02, Normal: 0x00, 0: 0, 1: 1, 2: 2}
   Static LVGS6 := {Collapsed: 0x01, Collapsible: 0x08, Focused: 0x10, Hidden: 0x02, NoHeader: 0x04, Normal: 0x00
                 , Selected: 0x20, 0: 0, 1: 1, 2: 2, 4: 4, 8: 8, 16: 16, 32: 32}
   Static LVGF := 0x04 ; LVGF_STATE
   Static SizeOfLVGROUP := (4 * 6) + (A_PtrSize * 4)
   Static OffStateMask := 8 + (A_PtrSize * 3) + 8
   Static OffState := OffStateMask + 4
   SetStates := 0
   LVGS := OS > 5 ? LVGS6 : LVGS5
   For Each, State In States {
      If !LVGS.HasKey(State)
         Return False
      SetStates |= LVGS[State]
   }
   VarSetCapacity(LVGROUP, SizeOfLVGROUP, 0)
   NumPut(SizeOfLVGROUP, LVGROUP, 0, "UInt")
   NumPut(LVGF, LVGROUP, 4, "UInt")
   NumPut(SetStates, LVGROUP, OffStateMask, "UInt")
   NumPut(SetStates, LVGROUP, OffState, "UInt")
   return SendMessage(0x1093, GroupID, &LVGROUP, , "ahk_id " . HLV)
}

; ======================================================================================================================
; LV_EX_GroupRemove - Removes a group from a list-view control.
; ======================================================================================================================
LV_EX_GroupRemove(HLV, GroupID) {
    ; LVM_REMOVEGROUP = 0x1096 -> msdn.microsoft.com/en-us/library/bb761149(v=vs.85).aspx
    return SendMessage(0x10A0, GroupID, 0, , "ahk_id " . HLV)
}

; ======================================================================================================================
; LV_EX_SetGroup - Assigns a list-view item to an existing group.
; ======================================================================================================================
LV_EX_SetGroup(HLV, Row, GroupID) {
   ; LVM_SETITEMA = 0x1006 -> http://msdn.microsoft.com/en-us/library/bb761186(v=vs.85).aspx
   Static OffGroupID := 28 + (A_PtrSize * 3)
   LV_EX_LVITEM(LVITEM, 0x00000100, Row) ; LVIF_GROUPID
   NumPut(GroupID, LVITEM, OffGroupID, "UPtr")
   return SendMessage(0x1006, 0, &LVITEM, , "ahk_id " . HLV)
}

; ======================================================================================================================
; ======================================================================================================================
; Function for internal use ============================================================================================
; ======================================================================================================================
; ======================================================================================================================
LV_EX_LVITEM(ByRef LVITEM, Mask := 0, Row := 1, Col := 1) {
   Static LVITEMSize := 48 + (A_PtrSize * 3)
   VarSetCapacity(LVITEM, LVITEMSize, 0)
   NumPut(Mask, LVITEM, 0, "UInt"), NumPut(Row - 1, LVITEM, 4, "Int"), NumPut(Col - 1, LVITEM, 8, "Int")
}
; ----------------------------------------------------------------------------------------------------------------------
LV_EX_PWSTR(Str, ByRef WSTR) { ; ANSI to Unicode
   VarSetCapacity(WSTR, StrPut(Str, "UTF-16") * 2, 0)
   StrPut(Str, &WSTR, "UTF-16")
   Return &WSTR
}

; ======================================================================================================================
; SetImage        Sets an image from the header's image list for the specified item.
; Parameters:     Image - 1-based index of the image in the image list.
; Return values:  Returns nonzero upon success, or zero otherwise.
; ======================================================================================================================
HD_EX_SetImage(HHD, Index, Image) {
   Static HDM_SETITEM := A_IsUnicode ? 0x120C : 0x1204 ; HDM_SETITEMW : HDM_SETITEMA
   Static HDF_IMAGE := 0x0800
   Static HDI_FORMAT := 0x0004
   Static HDI_IMAGE  := 0x0020
   Static OffFmt := (4 * 3) + (A_PtrSize * 2)
   Static OffImg := (4 * 4) + (A_PtrSize * 3)
   Mask := HDI_FORMAT | HDI_IMAGE
   Fmt := HD_EX_GetFormat(HHD, Index) | HDF_IMAGE
   HD_EX_CreateHDITEM(HDITEM)
   NumPut(Mask, HDITEM, 0, "UInt")
   NumPut(Fmt, HDITEM, OffFmt, "Int")
   NumPut(Image - 1, HDITEM, OffImg, "Int")
   return SendMessage(HDM_SETITEM, Index - 1, &HDITEM, , "ahk_id " . HHD)
}
; ======================================================================================================================
; SetImageList    Assigns an image list to a header control.
; Parameters:     HIL - Handle to the image list.
; Return values:  Returns 0 upon failure or if no image list was set previously; otherwise it returns the handle to
;                 the image list previously associated with the control.
; ======================================================================================================================
HD_EX_SetImageList(HHD, HIL) {
   Static HDM_SETIMAGELIST := 0x1208
   Static HDSIL_NORMAL := 0
   return SendMessage(HDM_SETIMAGELIST, HDSIL_NORMAL, HIL, , "ahk_id " . HHD)
}

; ======================================================================================================================
; SetFormat       Sets the format of the specified item.
; Parameters:     FormatArray - Array containing one ore more of the format strings defined in HDF.
;                 Exclusive   - If False, the passed format flags will be added using a bitwise-or operation.
;                               Otherwise, existing format flags will be reset.
; Return values:  Returns nonzero upon success, or zero otherwise.
; ======================================================================================================================
HD_EX_SetFormat(HHD, Index, FormatArray, Exclusive := False) {
   Static HDM_SETITEM := A_IsUnicode ? 0x120C : 0x1204 ; HDM_SETITEMW : HDM_SETITEMA
   Static HDF := {Left: 0x0000, Right: 0x0001, Center: 0x0002
                , Bitmap: 0x2000, BitmapOnRight: 0x1000, OwnerDraw: 0x8000, String: 0x4000
                , Image: 0x0800, RtlReading: 0x0004, SortDown: 0x0200, SortUp: 0x0400, SplitButton: 0x1000000}
   Static HDI_FORMAT := 0x0004
   Static OffFmt := (4 * 3) + (A_PtrSize * 2)
   Fmt := Exclusive ? 0 : HD_EX_GetFormat(HDD, Index)
   For Each, Format In FormatArray
      If HDF.HasKey(Format)
         Fmt |= HDF[Format]
   HD_EX_CreateHDITEM(HDITEM)
   NumPut(HDI_FORMAT, HDITEM, 0, "UInt")
   NumPut(Fmt, HDITEM, OffFmt, "UInt")
   return SendMessage(HDM_SETITEM, Index - 1, &HDITEM, , "ahk_id " . HHD)
}

; ======================================================================================================================
; GetFormat       Gets the format of the specified item.
; Return values:  Returns the current item format flags if successful, or 0 otherwise.
; ======================================================================================================================
HD_EX_GetFormat(HHD, Index) {
   Static HDM_GETITEM := A_IsUnicode ? 0x120B : 0x1203 ; HDM_GETITEMW : HDM_GETITEMA
   Static HDI_FORMAT := 0x0004
   Static OffFmt := (4 * 3) + (A_PtrSize * 2)
   HD_EX_CreateHDITEM(HDITEM)
   NumPut(HDI_FORMAT, HDITEM, 0, "UInt")
   SendMessage(HDM_GETITEM, Index - 1, &HDITEM, , "ahk_id " . HHD)
   Return NumGet(HDITEM, OffFmt, "Int")
}

; ======================================================================================================================
; CreateHDITEM    Creates a HDITEM structure - for internal use!!!
; ======================================================================================================================
HD_EX_CreateHDITEM(ByRef HDITEM) {
   Static cbHDITEM := (4 * 6) + (A_PtrSize * 6)
   VarSetCapacity(HDITEM, cbHDITEM, 0)
   Return True
}

; ======================================================================================================================
; GetCount        Gets the count of the items in a header control.
; Return values:  Returns the number of items if successful, or -1 otherwise.
; ======================================================================================================================
HD_EX_GetCount(HHD) {
   Static HDM_GETITEMCOUNT := 0x1200
   return SendMessage(HDM_GETITEMCOUNT, 0, 0, , "ahk_id " . HHD)
}

keywords(key := "") {
    static keywords := {
        keywords: "abort action add after all alter analyze and as asc attach autoincrement before begin between by cascade case cast check collate column commit conflict constraint create cross current current_date current_time current_timestamp database default deferrable deferred delete desc detach distinct do drop each else end escape except exclusive exists explain fail filter following for foreign from full glob group having if ignore immediate in index indexed initially inner insert instead intersect into is isnull join key left like limit match natural no not nothing notnull null of offset on or order outer over partition plan pragma preceding primary query raise range recursive references regexp reindex release rename replace restrict right rollback row rows savepoint select set table temp temporary then to transaction trigger unbounded union unique update using vacuum values view virtual when where window with without",
        functions: "abs avg changes char coalesce count cume_dist date datetime dense_rank first_value glob group_concat hex ifnull instr json json_array json_array_length json_extract json_insert json_object json_patch json_remove json_replace json_set json_type json_valid json_quote json_group_array json_group_object json_each json_tree julianday lag last_insert_rowid last_value lead length like likelihood likely load_extension lower ltrim max min nth_value ntile nullif percent_rank printf quote random randomblob rank replace round row_number rtrim soundex sqlite_compileoption_get sqlite_compileoption_used sqlite_offset sqlite_source_id sqlite_version strftime substr sum time total total_changes trim typeof unicode unlikely upper zeroblob"
    }
    
    return keywords.HasKey(key) ? keywords[key] : ""
}

/**
 * Sets the Explorer theme on ListViews or TreeViews
 * @param {HCTL}: handle of a ListView or TreeView control}
 * @return: True/False
*/
SetExplorerTheme(HCTL) {
    If (DllCall("GetVersion", "UChar") > 5) {
        VarSetCapacity(ClassName, 1024, 0)
        If DllCall("GetClassName", "Ptr", HCTL, "Str", ClassName, "Int", 512, "Int") {
            If (ClassName = "SysListView32") || (ClassName = "SysTreeView32")
                Return (!DllCall("UxTheme.dll\SetWindowTheme", "Ptr", HCTL, "WStr", "Explorer", "Ptr", 0))
        }
    }
    Return False
}

; ==================================================================================================================================
; Hides the focus border for the given GUI control or GUI and all of its children.
; Call the function passing only the HWND of the control / GUI in wParam as only parameter.
; WM_UPDATEUISTATE  -> msdn.microsoft.com/en-us/library/ms646361(v=vs.85).aspx
; The Old New Thing -> blogs.msdn.com/b/oldnewthing/archive/2013/05/16/10419105.aspx
; ==================================================================================================================================
HideFocusBorder(wParam, lParam := "", uMsg := "", hWnd := "") {
   ; WM_UPDATEUISTATE = 0x0128
	Static Affected := [] ; affected controls / GUIs
        , HideFocus := 0x00010001 ; UIS_SET << 16 | UISF_HIDEFOCUS
	     , OnMsg := OnMessage(0x0128, Func("HideFocusBorder"))
	If (uMsg = 0x0128) { ; called by OnMessage()
        If (wParam = HideFocus)
            Affected[hWnd] := True
        Else If Affected[hWnd]
            PostMessage(0x0128, HideFocus, 0, , "ahk_id " hWnd)
    }
    Else If DllCall("IsWindow", "Ptr", wParam, "UInt")
        PostMessage(0x0128, HideFocus, 0, , "ahk_id " wParam)
}

CvtClr(Color) {
    Return (Color & 0xFF) << 16 | (Color & 0xFF00) | (Color >> 16)
}
