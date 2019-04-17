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

main := GuiCreate("+MaximizeBox +Resize +MinSize800x600", "SQLite Viewer")
main.SetFont("s9", "Tahoma")
main.OnEvent("Close", (gui) => mainClose(gui, snippetsTab))
main.OnEvent("Size", "mainResize")

dbViewerTreeview := new SQLiteViewer_DBTreeview(main, 600, 200)
addQueryTabFn := addQueryTab(db, name) => (queryTabs.addQueryTab(db, name), resultsLV.resetEmptyText("Run a query!", true))
dbViewerTreeview.OnConnect(addQueryTabFn)
dbViewerTreeview.OnDisconnect((db) => queryTabs.removeQueryTab(db))
dbViewerTreeview.OnNewWorksheet(addQueryTabFn)

global queryTabs := new SQLiteViewer_QueryTabs(main, "h200 w700 ys -Wrap Theme Section vQueryTabs")

runQueryBtn := main.addButton("w" 100 - main.marginX " ys+" queryTabs.tabsInterior.y " vRunBtn Disabled", "Run")
runQueryBtn.OnEvent("Click", (ctrl) => (tabIndex := queryTabs.getActiveTab(), runQuery(queryTabs.tabs[tabIndex], queryTabs.getQueryEditText())))

clearResultsBtn := main.addButton("wp vClearBtn", "Clear Results")
clearResultsBtn.OnEvent("Click", () => (resultsLV.clearRows(true), resultsLV.resetEmptyText("Run a query!", true)))

utilitiesTabs := main.addTab3("ys w200 h600 vUtilities", "Snippets")
TC_EX_GetInterior(utilitiesTabs.hwnd, tabx, taby, tabw, tabh)

snippetsTab := new SQLiteViewer_Snippets(main, tabw - 2, tabh - taby, loadSnippets())
utilitiesTabs.UseTab()

splitter := main.addText("w800 h2 xs y" queryTabs.queryTabs.pos.h + main.MarginY * 2 " Section 0x8 vSplitter")
splitter.OnEvent("click", "splitterDrag")

resultsTabs := main.addTab3("w800 h" 600 - queryTabs.queryTabs.pos.h - (main.marginY * 2) - splitter.pos.h " xs vResultsTabs", "Results|History")
TC_EX_GetInterior(resultsTabs.hwnd, tabx, taby, tabw, tabh)
resultsTabs.useTab(1)

resultsLV := new SQLiteViewer_ResultsListView(main, tabh - taby, tabw - 2)
resultsLV.OnLinkClick := () => dbViewerTreeview.promptForDb()

resultsTabs.useTab(2)

historyLV := new SQLiteViewer_HistoryTab(main, tabh - taby - (main.MarginY * 2), (tabw - 2) // 2)

resultsTabs.UseTab()

statusBar := main.addStatusBar("vStatusbar")
statusBar.setParts(200, 600, 200)
main.show()

global verticalSplitCursor := DllCall("LoadCursor", "UInt", 0, "UInt", IDC_SIZENS := 32645)
DllCall("SetClassLongPtrW", "Uint", splitter.hwnd, "Int", GCL_HCURSOR := -12, "Ptr", verticalSplitCursor) ; Set the cursor for the splitter

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
    ; cb := (this, count, text, names) => MsgBox(count)
    ; stat := db.exec(sql)
    stat := db.Query(sql, RS)
    DllCall("QueryPerformanceCounter", "Int64*", queryAfter)
    queryTime := Round((queryAfter - queryBefore) / freq * 1000, 3)
    
    if (!stat) {
        historyLV.addRow(fileName, sql, "ERROR")
        SendMessage(SB_SETICON, 2, error_icon, statusBar.hwnd) ; Otherwise SB.SetIcon() destroys the previous icon
        statusBar.SetText("ERROR - " db.ErrorMsg, 3)
        if (historyLV.LV.GetCount() > LV_EX_GetRowsPerPage(historyLV.LV.hwnd)) {
            historyLV.LV.ModifyCol(2, "AutoHdr")
        }
        MsgBox("Code: " db.ErrorCode "`n`nMessage:`n" db.ErrorMsg, "Query Error!", 16)
        return
    }
    
    rowCount := resultsLV.setData(RS)

    historyLV.addRow(fileName, sql, rowCount)
    if (historyLV.LV.GetCount() > LV_EX_GetRowsPerPage(historyLV.LV.hwnd)) {
        historyLV.LV.ModifyCol(2, "AutoHdr")
    }
    
    SendMessage(SB_SETICON, 2, successful_icon, statusBar.hwnd) ; Otherwise SB.SetIcon() destroys the previous icon
    statusBar.setText(rowCount " rows", 3)
    statusBar.SetText(queryTime "ms elapsed", 4)
}

splitterDrag(ctrl) {
    gui := ctrl.gui
    
    ; change the cursor for the GUI as a whole while the splitter is being dragged to avoid it constantly changing between pointer and SIZENS
    oldCursor := DllCall("SetClassLongPtrW", "Uint", gui.hwnd, "Int", GCL_HCURSOR := -12, "Ptr", verticalSplitCursor)
    
    MouseGetPos(, initY)
    while (GetKeyState("LButton")) {
        MouseGetPos(, newY)
        diffY := newY - initY
        
        ; restrict vertical movement
        tooHigh := newY <= gui.control["ClearBtn"].pos.y + gui.control["ClearBtn"].pos.h + gui.marginY
        tooLow := newY >= gui.pos.h - gui.pos.h * 0.3
        if (tooHigh || tooLow) {
            continue
        }
        
        gui.control["splitter"].move("y" newY)
        
        gui.control["QueryTabs"].move("h" gui.control["QueryTabs"].pos.h + diffY)
        gui.control["Query"].move("h" gui.control["Query"].pos.h + diffY)
        
        gui.control["ResultsTabs"].move("y" gui.control["ResultsTabs"].pos.y + diffY " h" gui.control["ResultsTabs"].pos.h - diffY)
        
        ; Since these are in a Tab3, we don't even have to adjust the y position.  AHK handles it for us
        gui.control["ResultsLV"].move("h" gui.control["ResultsLV"].pos.h - diffY)
        gui.control["HistoryLV"].move("h" gui.control["HistoryLV"].pos.h - diffY)
        gui.control["HistoryEdit"].move("h" gui.control["HistoryEdit"].pos.h - diffY)
        
        ; reset the Y value to diff against for next iteration
        initY := newY
        sleep(40) ; a litte delay before moving again
    }
    ; change the GUI cursor back when done
    DllCall("SetClassLongPtrW", "Uint", gui.hwnd, "Int", GCL_HCURSOR := -12, "Ptr", oldCursor)
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
    if (buffer.count = 3 || minMax = 1 || (prevMinMax > 0 && minMax = 0)) {
        halfH := buffer.h // 2
        remainderH := Mod(buffer.h, 2)
        halfW := buffer.w // 2
        remainderW := Mod(buffer.w, 2)
        
        gui.control["DBtreeview"].move("h" gui.control["DBtreeview"].pos.h + buffer.h)
        gui.control["QueryTabs"].move("h" gui.control["QueryTabs"].pos.h + halfH " w" gui.control["QueryTabs"].pos.w + buffer.w)
        gui.control["Query"].move("h" gui.control["Query"].pos.h + halfH " w" gui.control["Query"].pos.w + buffer.w)
        gui.control["ResultsTabs"].move("y" gui.control["ResultsTabs"].pos.y + halfH " h" gui.control["ResultsTabs"].pos.h + halfH + remainderH " w" gui.control["ResultsTabs"].pos.w + buffer.w)
        gui.control["splitter"].move("y" (gui.marginY * 2) + gui.control["QueryTabs"].pos.h " w" gui.control["ResultsTabs"].pos.w, true)
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
    sci.AutoCSetFillups("", "([.", 1)
    
    ; Function icon used in AutoComplete - must set dimensions first
    sci.RGBAImageSetWidth(16)
    sci.RGBAImageSetHeight(16)
    sci.RegisterRGBAImage(1, &function_icon := getFunctionIcon())
    
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
    
    ; Indicators
    sci.IndicSetStyle(8, 0)
    sci.IndicSetFore(8, 0xF8F8F2)

    ; sci.StyleSetFore(sci.STYLE_BRACELIGHT, CvtClr(0x3399FF))
    ; sci.StyleSetBold(sci.STYLE_BRACELIGHT, True)

    sci.StyleSetFore(sci.SCE_SQL_COMMENT, CvtClr(0x75715E))
    sci.StyleSetFore(sci.SCE_SQL_COMMENTLINE, CvtClr(0x75715E))
    sci.StyleSetFore(sci.SCE_SQL_COMMENTDOC, CvtClr(0x75715E))
    sci.StyleSetFore(sci.SCE_SQL_COMMENTDOCKEYWORD, CvtClr(0x66D9EF))
    sci.StyleSetFore(sci.SCE_SQL_WORD, CvtClr(0xF92672))
    sci.StyleSetBold(sci.SCE_SQL_WORD, false)
    sci.StyleSetFore(sci.SCE_SQL_NUMBER, CvtClr(0xAE81FF))
    sci.StyleSetFore(sci.SCE_SQL_STRING, CvtClr(0xE6DB74))
    sci.StyleSetFore(sci.SCE_SQL_CHARACTER, CvtClr(0xE6DB74)) ; single quoted strings
    sci.StyleSetFore(sci.SCE_SQL_OPERATOR, CvtClr(0xF92672))
    sci.StyleSetFore(sci.SCE_SQL_USER1, CvtClr(0x66D9EF))

    sci.SetKeywords(0, keywords("keywords"), 1)
    sci.SetKeywords(4, StrReplace(keywords("functions"), "?1"), 1) ; remove the icon identifiers so they get highlighted properly by the lexer

    ; line number margin
    PixelWidth := sci.TextWidth(sci.STYLE_LINENUMBER, "9999", 1)
    sci.SetMarginWidthN(0, PixelWidth)
    sci.SetMarginLeft(0, 2) ; Left padding
    
    ; used as a border between line numbers and content
    borderMarginW := 0
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
    VarSetCapacity(LVGROUP, SizeOfLVGROUP, 0)
    NumPut(SizeOfLVGROUP, LVGROUP, 0, "UInt")
    NumPut(Mask, LVGROUP, 4, "UInt")
    NumPut(&Header, LVGROUP, OffHeader, "Ptr")
    NumPut(GroupID, LVGROUP, OffGroupID, "Int")
    NumPut(Align, LVGROUP, OffAlign, "UInt")
    NumPut(&Subtitle, LVGROUP, OffSubtitle, "Ptr")
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
                , Bitmap: 0x2000, BitmapOnRight: 0x1000, OwnerDraw: 0x8000, String: 0x4000, Checkbox: 0x40, Checked: 0x80
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

getfunctionIcon() {
    static functionIconSrc := "0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xe4|0xe4|0xe4|0xff|0xa5|0xa5|0xa5|0xff|0xa9|0xa9|0xa9|0xff|0xf8|0xf8|0xf8|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xa4|0xa4|0xa4|0xff|0x2a|0x2a|0x2a|0xff|0x2c|0x2c|0x2c|0xff|0x0|0x0|0x0|0xff|0x7b|0x7b|0x7b|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xe3|0xe3|0xe3|0xff|0x9|0x9|0x9|0xff|0x7c|0x7c|0x7c|0xff|0xbf|0xbf|0xbf|0xff|0x3e|0x3e|0x3e|0xff|0xbe|0xbe|0xbe|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x8a|0x8a|0x8a|0xff|0x0|0x0|0x0|0xff|0xa8|0xa8|0xa8|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x41|0x41|0x41|0xff|0x0|0x0|0x0|0xff|0xcd|0xcd|0xcd|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x89|0x89|0x89|0xff|0x55|0x55|0x55|0xff|0xb|0xb|0xb|0xff|0x0|0x0|0x0|0xff|0x4c|0x4c|0x4c|0xff|0x6b|0x6b|0x6b|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xae|0xae|0xae|0xff|0x83|0x83|0x83|0xff|0x0|0x0|0x0|0xff|0x16|0x16|0x16|0xff|0x99|0x99|0x99|0xff|0xb2|0xb2|0xb2|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xb5|0xb5|0xb5|0xff|0x0|0x0|0x0|0xff|0x48|0x48|0x48|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xf2|0xf2|0xf2|0xff|0x95|0x95|0x95|0xff|0x7c|0x7c|0x7c|0xff|0xec|0xec|0xec|0xff|0xe5|0xe5|0xe5|0xff|0x7f|0x7f|0x7f|0xff|0xf1|0xf1|0xf1|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x85|0x85|0x85|0xff|0x0|0x0|0x0|0xff|0x75|0x75|0x75|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xe4|0xe4|0xe4|0xff|0xb0|0xb0|0xb0|0xff|0x1|0x1|0x1|0xff|0x70|0x70|0x70|0xff|0x38|0x38|0x38|0xff|0x4|0x4|0x4|0xff|0xd8|0xd8|0xd8|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x55|0x55|0x55|0xff|0x0|0x0|0x0|0xff|0xaf|0xaf|0xaf|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x3b|0x3b|0x3b|0xff|0x10|0x10|0x10|0xff|0xe5|0xe5|0xe5|0xff|0xf1|0xf1|0xf1|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xfd|0xfd|0xfd|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x18|0x18|0x18|0xff|0xc|0xc|0xc|0xff|0xf3|0xf3|0xf3|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xfe|0xfe|0xfe|0xff|0xfc|0xfc|0xfc|0xff|0x59|0x59|0x59|0xff|0x1|0x1|0x1|0xff|0xe5|0xe5|0xe5|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xb9|0xb9|0xb9|0xff|0x6|0x6|0x6|0xff|0x84|0x84|0x84|0xff|0xd0|0xd0|0xd0|0xff|0x0|0x0|0x0|0xff|0x7f|0x7f|0x7f|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x45|0x45|0x45|0xff|0xf|0xf|0xf|0xff|0x8f|0x8f|0x8f|0xff|0x0|0x0|0x0|0xff|0x89|0x89|0x89|0xff|0xec|0xec|0xec|0xff|0xff|0xff|0xff|0xff|0xb4|0xb4|0xb4|0xff|0x2|0x2|0x2|0xff|0xe|0xe|0xe|0xff|0x48|0x48|0x48|0xff|0x76|0x76|0x76|0xff|0xfc|0xfc|0xfc|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0x86|0x86|0x86|0xff|0xa1|0xa1|0xa1|0xff|0xfd|0xfd|0xfd|0xff|0x82|0x82|0x82|0xff|0x63|0x63|0x63|0xff|0xba|0xba|0xba|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xe4|0xe4|0xe4|0xff|0xd5|0xd5|0xd5|0xff|0xf6|0xf6|0xf6|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff|0xff"
    
    ; 16x16 image with 4 bytes per pixel
    VarSetCapacity(functionIcon, 4 * 16 * 16 + 1)
    Loop parse, functionIconSrc, "|" {
        NumPut(A_LoopField, functionIcon, A_Index - 1, "UInt")
    }
    
    return functionIcon
}

keywords(key := "") {
    static keywords := {
        keywords: "abort action add after all alter analyze and as asc attach autoincrement before begin between by cascade case cast check collate column commit conflict constraint create cross current current_date current_time current_timestamp database default deferrable deferred delete desc detach distinct do drop each else end escape except exclusive exists explain fail filter following for foreign from full glob group having if ignore immediate in index indexed initially inner insert instead intersect into is isnull join key left like limit match natural no not nothing notnull null of offset on or order outer over partition plan pragma preceding primary query raise range recursive references regexp reindex release rename replace restrict right rollback row rows savepoint select set table temp temporary then to transaction trigger unbounded union unique update using vacuum values view virtual when where window with without",
        functions: "abs?1 avg?1 changes?1 char?1 coalesce?1 count?1 cume_dist?1 date?1 datetime?1 dense_rank?1 first_value?1 glob?1 group_concat?1 hex?1 ifnull?1 instr?1 json?1 json_array?1 json_array_length?1 json_extract?1 json_insert?1 json_object?1 json_patch?1 json_remove?1 json_replace?1 json_set?1 json_type?1 json_valid?1 json_quote?1 json_group_array?1 json_group_object?1 json_each?1 json_tree?1 julianday?1 lag?1 last_insert_rowid?1 last_value?1 lead?1 length?1 like?1 likelihood?1 likely?1 load_extension?1 lower?1 ltrim?1 max?1 min?1 nth_value?1 ntile?1 nullif?1 percent_rank?1 printf?1 quote?1 random?1 randomblob?1 rank?1 replace?1 round?1 row_number?1 rtrim?1 soundex?1 sqlite_compileoption_get?1 sqlite_compileoption_used?1 sqlite_offset?1 sqlite_source_id?1 sqlite_version?1 strftime?1 substr?1 sum?1 time?1 total?1 total_changes?1 trim?1 typeof?1 unicode?1 unlikely?1 upper?1 zeroblob?1"
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
