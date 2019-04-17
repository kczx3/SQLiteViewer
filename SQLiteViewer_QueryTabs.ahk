class SQLiteViewer_QueryTabs {
    tabs := []
    tabContents := []
    lastTab := 1
    calltipHighlightPos := 0
    currCalltip := false
    prevPosition := 0
    
    __New(gui, opts) {
        this.gui := Gui
        
        this.queryTabs := gui.addTab2(Opts, "+")
        this.queryTabs.OnEvent("Change", (ctrl) => this.onTabChange(ctrl))
        this.queryTabs.OnNotify(TCN_SELCHANGING := -552, (ctrl) => this.onTabChanging(ctrl))
        OnMessage(WM_MBUTTONDOWN  := 0x0207, (wParam, lParam, msg, hwnd) => this.OnMClick(wParam, lParam, msg, hwnd))
        this.queryTabs.UseTab()
        TC_EX_GetInterior(this.queryTabs.hwnd, tabx, taby, tabw, tabh)
        this.tabsInterior := {
            x: tabx,
            y: taby,
            w: tabw,
            h: tabh
        }
        
        this.queryEdit := new Scintilla(gui, "x" this.queryTabs.pos.x + this.tabsInterior.x - 2 " y" this.queryTabs.pos.y + this.tabsInterior.y " w" this.tabsInterior.w - 2 " h" this.tabsInterior.h - this.tabsInterior.y + 1 " Hidden vQuery", , 0, 0)
        setupSciControl(this.queryEdit)
        
        this.queryEdit.SetMouseDwellTime(2000)
        this.queryEdit.SetModEventMask(Scintilla.SC_MOD_BEFOREDELETE | Scintilla.SC_PERFORMED_USER | Scintilla.SC_MOD_CHANGESTYLE | Scintilla.SC_MOD_INSERTCHECK)
        this.queryEdit.OnNotify(this.queryEdit.SCN_CHARADDED, (ctrl, l) => this.handleCharAdded(ctrl, l))
        this.queryEdit.OnNotify(this.queryEdit.SCN_MODIFIED, (ctrl, l) => this.handleModification(ctrl, l))
        this.queryEdit.OnNotify(this.queryEdit.SCN_AUTOCSELECTION, (ctrl, l) => this.handleAutoCSelection(ctrl, l))
        this.queryEdit.OnNotify(this.queryEdit.SCN_AUTOCCOMPLETED, (ctrl, l) => this.handleAutoCCompleted(ctrl, l))
        this.queryEdit.OnNotify(this.queryEdit.SCN_UPDATEUI, (ctrl, l) => this.handleUpdateUI(ctrl, l))
        this.queryEdit.OnNotify(this.queryEdit.SCN_DWELLSTART, (ctrl, l) => this.handleDwell(ctrl, l))
        this.queryEdit.OnNotify(this.queryEdit.SCN_DWELLEND, (ctrl, l) => this.handleDwell(ctrl, l, true))
    }
    
    OnMClick(wParam, lParam, msg, hwnd) {
        if (hwnd = this.queryTabs.hwnd) {
            VarSetCapacity(HITTEST, 12)
            NumPut(lParam & 0xFFFF, HITTEST, 0, "Int")
            NumPut(lParam >> 16, HITTEST, 4, "Int")
            index := SendMessage(TCM_HITTEST := 0x130D, 0, &HITTEST, hwnd)
            this.removeQueryTab(false, index + 1)
        }
    }
    
    handleAutoCSelection(ctrl, l) {
        ; only process for autocompletion done via TAB and ENTER
        if ((isTab := this.queryEdit.listCompletionMethod = this.queryEdit.SC_AC_TAB) || this.queryEdit.listCompletionMethod = this.queryEdit.SC_AC_NEWLINE) {
            ; get what the user had already typed
            start := this.queryEdit.position
            end := this.queryEdit.WordEndPosition(start)
            word := GetTextRange([start, end])
            
            ; if it matches a whole word from the autocompletion list
            if (word = StrGet(this.queryEdit.text, "UTF-8")) {
                ; prep to insert accordingly upon receipt of SCN_AUTOCCOMPLETED
                if (isTab) {
                    this.insertAfterAutoComplete := "    "
                }
                else {
                    this.insertAfterAutoComplete := "`n"
                }
            }
        }
        
        ; helper
        GetTextRange(Range) {
            VarSetCapacity(Text, Abs(Range[1] - Range[2]) + 1, 0)
            VarSetCapacity(Sci_TextRange, 8 + A_PtrSize, 0)
            NumPut(Range[1], Sci_TextRange, 0, "UInt")
            NumPut(Range[2], Sci_TextRange, 4, "UInt")
            NumPut(&Text, Sci_TextRange, 8, "Ptr")
            this.queryEdit.GetTextRange(0, &Sci_TextRange) ; SCI_GETTEXTRANGE
            Return StrGet(&Text,, "UTF-8")
        }
    }
    
    handleAutoCCompleted(ctrl, l) {
        if (this.insertAfterAutoComplete) {
            this.queryEdit.InsertText(caretPos := this.queryEdit.GetCurrentPos(), this.insertAfterAutoComplete, 1)
            this.queryEdit.GoToPos(caretPos + StrLen(this.insertAfterAutoComplete))
            this.insertAfterAutoComplete := ""
        }
    }
    
    /**
     * Need to handle commas while CallTips are active here because we need to check the style of the comma
     * Only increment/decrement CallTip highlighting if the comma is of style SCE_SQL_OPERATOR
     */
    handleModification(ctrl, l) {
        ; Before deletion and invoked by the user
        ; allows us to adjust calltip highlighting when user is deleting text inside a functions parameters
        if (this.queryEdit.modType & Scintilla.SC_MOD_BEFOREDELETE && this.queryEdit.modType & Scintilla.SC_PERFORMED_USER) {
            ; only update the highlighted calltip parameter if they are removing a single character
            ; Not sure how to handle this otherwise
            if (this.queryEdit.length = 1) {
                this.alterCalltipHighlight(false)
            }
            
            caretPos := this.queryEdit.Position
            docStart := caretPos = 1
            docEnd := caretPos = this.queryEdit.GetLength()
            prevChar := Chr(this.queryEdit.GetCharAt(caretPos))
            nextChar := Chr(this.queryEdit.GetCharAt(caretPos + 1))
            
            ; store a boolean to determine if we should delete the character to the right of the cursor after this deletion
            ; this is done in response to SCN_UPDATEUI
            this.shouldDeleteNext := (prevChar = "`"" && nextChar = "`"") || (prevChar = "'" && nextChar = "'") || (prevChar = "[" && nextChar = "]") || (prevChar = "{" && nextChar = "}") || (prevChar = "(" && nextChar = ")")
        }
        ; Style change - used for when commas are typed to adjust calltip highlighting
        else if (this.queryEdit.modType & Scintilla.SC_MOD_CHANGESTYLE) {
            this.alterCalltipHighlight(true)
        }
        
        ; else if (this.queryEdit.modType & Scintilla.SC_MOD_INSERTCHECK) {
            ; toInsert := StrGet(this.queryEdit.text, , "UTF-8")            ; if (toInsert = "`r`n") {
                ; caretPos := this.queryEdit.Position
                ; prevChar := Chr(this.queryEdit.GetCharAt(caretPos - 1))
                ; nextChar := Chr(this.queryEdit.GetCharAt(caretPos))
                ; 
                ; autoIndent := prevChar = "{" && nextChar = "}"
                ; 
                ; if (autoIndent) {
                    ; Line := this.queryEdit.LineFromPosition(caretPos)
                    ; iIndentation := this.queryEdit.GetLineIndentation(Line)
                    ; sIndentation := iIndentation ? Format("{1: " . iIndentation . "}", "") : ""
                    ; 
                    ; this.queryEdit.ChangeInsertion(StrLen("`r`n" . sIndentation . "    " . "`r`n" . sIndentation), "`r`n" . sIndentation . "    " . "`r`n" . sIndentation, 1)
                    ; this.queryEdit.GoToPos(caretPos + StrLen(CRLF . sIndentation . "    "))
                ; }
            ; }
        ; }
        return
    }
    
    handleUpdateUI(ctrl, l) {
        ; selection is being updated
        if (this.queryEdit.updated & this.queryEdit.SC_UPDATE_SELECTION) {
            caretPos := this.queryEdit.GetCurrentPos()
            if (this.shouldDeleteNext) {
                this.queryEdit.DeleteRange(caretPos, 1)
                this.shouldDeleteNext := false
            }
            
            if (this.queryEdit.CallTipActive()) {
                caretPos := this.queryEdit.GetCurrentPos()
                currLine := this.queryEdit.LineFromPosition(caretPos)
                callTipPos := this.queryEdit.CallTipPosStart()
                callTipLine := this.queryEdit.LineFromPosition(callTipPos)
                selectionLength := Abs(this.queryEdit.GetSelectionStart() - this.queryEdit.GetSelectionEnd())
                
                charRight := Chr(this.queryEdit.GetCharAt(caretPos))
                styleRight := this.queryEdit.GetStyleAt(caretPos)
                charLeft := Chr(this.queryEdit.GetCharAt(caretPos - 1))
                styleLeft := this.queryEdit.GetStyleAt(caretPos - 1)
                
                
                ; reset calltip tracking if cursor moves before calltip start or to a different line than the calltip is for or they select more than 1 character on that line
                if (caretPos < callTipPos || currLine != callTipLine || selectionLength > 1) {
                    this.queryEdit.CallTipCancel()
                    this.calltipHighlightPos := 0
                    this.currCalltip := {}
                }
                
                ; character to the left is a comma and the caret was moved to the right
                if (charLeft = "," && styleLeft = this.queryEdit.SCE_SQL_OPERATOR && this.prevPosition < caretPos) {
                    this.alterCalltipHighlight(true, caretPos - 1) ; highlight next parameter
                }
                ; character to the right is a comma and the caret was moved to the left
                else if (charRight = "," && styleRight = this.queryEdit.SCE_SQL_OPERATOR && this.prevPosition > caretPos) {
                    this.alterCalltipHighlight(false, caretPos) ; highlight previous parameter
                }
                
                ; store the current position to be used in the future
                this.prevPosition := caretPos
            }
        }
    }
    
    handleDwell(ctrl, l, cancel := false) {
        if (cancel) {
            start := this.queryEdit.WordStartPosition(this.queryEdit.position)
            end := this.queryEdit.WordEndPosition(this.queryEdit.position)
            text := GetTextRange([start, end])
            this.queryEdit.CallTipCancel()
            this.queryEdit.SetIndicatorCurrent(8)
            this.queryEdit.IndicatorClearRange(start, end - start)
        }
        else {
            this.calltip(this.queryEdit.position, true)
        }
        
        ; helper
        GetTextRange(Range) {
            VarSetCapacity(Text, Abs(Range[1] - Range[2]) + 1, 0)
            VarSetCapacity(Sci_TextRange, 8 + A_PtrSize, 0)
            NumPut(Range[1], Sci_TextRange, 0, "UInt")
            NumPut(Range[2], Sci_TextRange, 4, "UInt")
            NumPut(&Text, Sci_TextRange, 8, "Ptr")
            this.queryEdit.GetTextRange(0, &Sci_TextRange) ; SCI_GETTEXTRANGE
            Return StrGet(&Text,, "UTF-8")
        }
    }
    
    handleCharAdded(ctrl, l) {
        static wordChars := "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        char := Chr(this.queryEdit.ch)
        caretPos := this.queryEdit.GetCurrentPos()
        
        ; LIMITATION: this would still cause a calltip to display inside of a string literal;
        ; Honestly, these would have to be handled within SCN_MODIFIED probably to properly handle the style that is applied to the parenthesis
        if (char ~= "\(|\)|\[|\]|\{|\}|`"|'") {
            this.autoClose(char, caretPos)
            
            if (char = "(") {
                this.callTip(caretPos)
            }
        }
        
        ; cancel the calltip because they typed a close parenthesis
        else if (char = ")") {
            if (this.queryEdit.CallTipActive()) {
                this.queryEdit.CallTipCancel()
                this.calltipHighlightPos := 0
                this.currCalltip := {}
            }
        }
        else if (char = "`n") {
            this.autoIndent(caretPos)
        }
        
        ; Autocomplete
        else if (InStr(wordChars, char) && !this.queryEdit.CallTipActive()) {
            this.autoComplete(caretPos)
        }
    }
    
    autoComplete(caretPos) {
        start := this.queryEdit.WordStartPosition(caretPos, 1)
        lenEntered := caretPos - start
        
        if (lenEntered > 2) {
            if (!this.queryEdit.AutoCActive()) {
                this.queryEdit.AutoCShow(lenEntered, keywords("keywords") " " keywords("functions"), 1)
            }
        }
    }
    
    autoIndent(caretPos) {
        prevChar := Chr(this.queryEdit.GetCharAt(caretPos - 3))
        nextChar := Chr(this.queryEdit.GetCharAt(caretPos))
        line := this.queryEdit.LineFromPosition(caretPos) - 1
        iIndentation := this.queryEdit.GetLineIndentation(line)
        sIndentation := iIndentation ? Format("{1: " . iIndentation . "}", "") : ""
        
        if (prevChar = "{" && nextChar = "}") {
            this.queryEdit.InsertText(caretPos, sIndentation . "    `r`n" . sIndentation, 1)
            this.queryEdit.GoToPos(caretPos + iIndentation + 4)
        }
        else if (prevChar = "{" && nextChar = "`r") {
            this.queryEdit.InsertText(caretPos, sIndentation . "    ", 1)
            this.queryEdit.GoToPos(caretPos + iIndentation + 4)
        }
        else if (this.queryEdit.ch = 10) {
            this.queryEdit.InsertText(caretPos, sIndentation, 1)
            this.queryEdit.GoToPos(caretPos + iIndentation)
        }
    }
    
    callTip(caretPos, indicator := false) {
        pos := caretPos - 2
        ; Get word that is before the "("
        start := this.queryEdit.WordStartPosition(pos)
        end := this.queryEdit.WordEndPosition(pos)
        text := GetTextRange([start, end])
        
        ; The word isn't defined in our calltip JSON
        if (!tip := getCallTip(text)) {
            return
        }
        
        ; store the calltip that we're showing currently
        this.currCalltip := tip
        
        ; no calltip shown so display it at the start position of the word we are showing the calltip for
        if (!this.queryEdit.CallTipActive()) {
            this.queryEdit.CallTipShow(start, tip.text, 1)
            
            if (indicator) {
                this.queryEdit.SetIndicatorCurrent(8)
                this.queryEdit.IndicatorFillRange(start, end - start)
            }
            
            ; apply highlighting if the word is configured for it
            if (tip.HasKey("highlight") && tip.highlight.length()) {
                this.queryEdit.CallTipSetHlt(tip.highlight[1][1], tip.highlight[1][2])
                this.calltipHighlightPos := 1 ; save which position in the highlight array we are currently highlighting
            }
        }
        
        ; helper
        GetTextRange(Range) {
            VarSetCapacity(Text, Abs(Range[1] - Range[2]) + 1, 0)
            VarSetCapacity(Sci_TextRange, 8 + A_PtrSize, 0)
            NumPut(Range[1], Sci_TextRange, 0, "UInt")
            NumPut(Range[2], Sci_TextRange, 4, "UInt")
            NumPut(&Text, Sci_TextRange, 8, "Ptr")
            this.queryEdit.GetTextRange(0, &Sci_TextRange) ; SCI_GETTEXTRANGE
            Return StrGet(&Text,, "UTF-8")
        }
    }
    
    autoClose(char, caretPos) {
        static matches := {")": "(", "}": "{", "]": "[", "`"": "`"", "'": "'"}
        
        docStart := caretPos = 1
        docEnd := caretPos = this.queryEdit.GetLength()
        
        prevChar := Chr(this.queryEdit.GetCharAt(docStart ? caretPos : caretPos - 2))
        nextChar := Chr(this.queryEdit.GetCharAt(caretPos))
        
        isPrevCharBlank := prevChar ~= "\s"
        isNextCharBlank := nextChar ~= "\s" || docEnd
        
        isEnclosed := (prevChar = "(" && nextChar = ")") || (prevChar = "{" && nextChar = "}") || (prevChar = "[" && nextChar = "}")
        isSpaceEnclosed := (prevChar == "(" && isNextCharBlank) || (isPrevCharBlank && nextChar == ")") || (prevChar == "{" && isNextCharBlank) || (isPrevCharBlank && nextChar == "}") || (prevChar == "[" && isNextCharBlank) || (isPrevCharBlank && nextChar == "]")
        
        isCharOrString := (isPrevCharBlank && isNextCharBlank) || isEnclosed || isSpaceEnclosed
        
        charNextIsCharOrString := nextChar = "`"" || charNext = "'"
        
        ; deletes the closing pair if you type it again while the pair is empty
        ; Example: typing ( will autoclose it and place the cursor in the middle. Then typing ) will simply delete it and move your cursor outside the pair
        if (prevChar = matches[char] && nextChar = char) {
            this.queryEdit.DeleteRange(caretPos, 1)
            this.queryEdit.GoToPos(caretPos)
        }
        
        ; close the character pair
        if (char = '"') {
            if (isCharOrString) {
                this.queryEdit.InsertText(caretPos, '"', 1)
            }
        }
        else if (char = "'") {
            if (isCharOrString) {
                this.queryEdit.InsertText(caretPos, "'", 1)
            }
        }
        else if (char = "(" && isNextCharBlank) {
            if (!charNextIsCharOrString) {
                this.queryEdit.InsertText(caretPos, ")", 1)
            }
        }
        else if (char = "{") {
            if (!charNextIsCharOrString) {
                this.queryEdit.InsertText(caretPos, "}", 1)
            }
        }
        else if (char = "[") {
            if (!charNextIsCharOrString) {
                this.queryEdit.InsertText(caretPos, "]", 1)
            }
        }
    }
    
    ; used to alter the highlight position of the calltip
    alterCalltipHighlight(increment, position := 0) {
        if (this.queryEdit.CallTipActive()) {
            position := position || this.queryEdit.position
            char := Chr(this.queryEdit.GetCharAt(position))
            
            if (char = "," && this.queryEdit.GetStyleAt(position) = this.queryEdit.SCE_SQL_OPERATOR) {
                if (increment) {
                    ++this.calltipHighlightPos
                }
                else {
                    --this.calltipHighlightPos
                }
                
                ; We increment the calltip highlight index no matter what, so we need to check if that index is defined for this calltip or not
                if (this.currCalltip.highlight.HasKey(this.calltipHighlightPos)) {
                    this.queryEdit.CallTipSetHlt(this.currCalltip.highlight[this.calltipHighlightPos][1], this.currCalltip.highlight[this.calltipHighlightPos][2])
                }
                ; if it isn't, then just clear the calltip highlighting
                ; the proper argument is re-highlighted if they backspace into an index that does have a highlight configured
                else {
                    this.queryEdit.CallTipSetHlt(0, 0)
                }
            }
        }
    }
    
    getQueryEditText() {
        len := this.queryEdit.GetLength() + 1
        VarSetCapacity(SciText, len, 0)
        this.queryEdit.GetText(len, &SciText)
        return StrGet(&SciText, "UTF-8")
    }
    
    addQueryTab(db, name) {
        newIndex := TC_EX_Add(this.queryTabs.hwnd, name, TC_EX_GetCount(this.queryTabs.hwnd) - 1)
        
        this.tabs.push(db)
        
        this.OnTabChanging(this.queryTabs)
        this.queryTabs.Choose(newIndex)
        
        if (!this.queryEdit.ctrl.visible) {
            this.queryEdit.ctrl.visible := true
            this.gui.control["runBtn"].Enabled := true
        }
        
        ; this.queryEdit.AddDocument()
        ; setupSciControl(this.queryEdit)
        this.queryEdit.ClearAll()
        this.queryEdit.ctrl.focus()
    }
    
    removeQueryTab(db, index := false) {
        if (!index) {
            Loop(length := this.tabs.length()) {
                i := length - A_Index + 1
                if (db == this.tabs[i]) {
                    this.queryTabs.Delete(i)
                    this.tabs.RemoveAt(i)
                    this.tabContents.RemoveAt(i)
                    ; this.queryEdit.deleteDocument(i)
                }
            }
        }
        else if (Type(index) = "Integer" && this.tabs.HasKey(index)) {
            this.queryTabs.Delete(index)
            this.tabs.RemoveAt(index)
            this.tabContents.RemoveAt(index)
        }
        
        this.queryTabs.Choose(index && Type(index) = "Integer" ? index : 1)
        this.onTabChange(this.queryTabs)
        ; this.queryEdit.switchDocument(1)
        
        if (TC_EX_GetCount(this.queryTabs.hwnd) = 1) {
            this.queryEdit.ctrl.visible := false
        }
    }
    
    getActiveTab() {
        return this.queryTabs.value
    }
    
    onTabChanging(ctrl) {
        if (ctrl.value != TC_EX_GetCount(ctrl.hwnd)) {
            start := this.queryEdit.GetSelectionStart()
            end := this.queryEdit.GetSelectionEnd()
            content := this.getQueryEditText()
            
            this.tabContents[ctrl.value] := {
                content: content,
                selection: [start, end]
            }
        }
    }
    
    onTabChange(ctrl) {
        if (ctrl.value != TC_EX_GetCount(ctrl.hwnd)) {
            this.queryEdit.ctrl.visible := true
            this.gui.control["runBtn"].Enabled := true
            if (this.tabContents.HasKey(ctrl.value)) {
                this.queryEdit.SetText("", this.tabContents[ctrl.value].content, 1)
                this.queryEdit.ctrl.focus()
                this.queryEdit.SetSel(this.tabContents[ctrl.value].selection[1], this.tabContents[ctrl.value].selection[2])
            }
            else {
                this.queryEdit.ClearAll()
                this.queryEdit.ctrl.focus()
            }
            ; this.queryEdit.switchDocument(ctrl.value)
        }
        else {
            this.queryEdit.ctrl.visible := false
            this.gui.control["runBtn"].Enabled := false
        }
    }
}