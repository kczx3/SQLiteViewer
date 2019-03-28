class SQLiteViewer_Snippets {
    __New(gui, width, height, snippets) {
        static toolbarCommands := {
            1: () => this.snippetGui(),
            2: () => this.snippetGui(true),
            3: () => this.removeSnippet()
        }
        this.gui := gui
        this.snippets := snippets || []
        
        this.toolbarImgList := IL_Create(1, 1, 0)
        IL_Add(this.toolbarImgList, "add.ico")
        IL_Add(this.toolbarImgList, "modify.ico")
        IL_Add(this.toolbarImgList, "delete.ico")
        
        this.tb := new Toolbar(this.gui, "x+0 y+0 h24 w" width " Menu ToolTips Transparent vSnippetsTB")
        this.tb.OnEvent("Click", (tb, id) => toolbarCommands[id]())
        this.tb.SetImageList(this.toolbarImgList)
        this.tb.SetMaxTextRows()
        this.tb.Add(, "Add", 0, , , , 1) ; Add Button
        this.tb.Add(, "Modify", 1, , , , 2) ; Modify Button
        this.tb.Add(, "Delete", 2, , , , 3) ; Delete Button
        this.tb.SetButtonSize(24, 24)
        this.tb.AutoSize()
        
        prevMarginy := this.gui.MarginY, this.gui.MarginY := 0
        this.divider := this.gui.addText("h1 w" width " 0x10")

        this.snippetsLV := this.gui.addListView("w" width " h" height - this.tb.ctrl.pos.h " -Hdr -E0x200 LV0x400 LV0x4000 0x100 Multi vSnippetsLV", "Snippets")
        
        this.loadSnippets()
        
        ; if DoubleClick is on an item, set that item's "content" as the query edit's value
        this.snippetsLV.OnEvent("DoubleClick", (ctrl, i) => i > 0 && ctrl.gui.control["query"].visible && queryTabs.queryEdit.SetText("", this.snippets[i].content, 1))
        
        this.snippetsLVTooltipHwnd := SendMessage(LVM_GETTOOLTIPS := 0x104E, 0, 0, this.snippetsLV.hwnd)
        this.snippetsLV.OnNotify(-158, (ctrl, l) => this.handleInfoTip(ctrl, l))
        
        SetExplorerTheme(this.snippetsLV.hwnd)
        HideFocusBorder(this.snippetsLV.hwnd)
        
        ; Make the tiles only have a single line (they always display one line for the first column's text) and set the width
        ; LV_EX_SetTileViewLines(this.snippetsLV.hwnd, 0, width - SysGet(2)) ; Subtract the width of the vertical scrollbar
        
        ; By applying LVS_AUTOARRANGE when creating the ListView and then setting the TileViewInfo after, we have to send LVM_UPDATE once to properly space the Tiles
        ; SendMessage(LVM_UPDATE := 0x102A, 0, 0, , "ahk_id " this.snippetsLV.hwnd)
        
        ; Assing the previous MarginY back to the GUI
        this.gui.MarginY := prevMarginY
    }
    
    handleInfoTip(ctrl, l) {
        Static NMHDRSize := A_PtrSize * 3
        Static offText := NMHDRSize + A_PtrSize
        Static offItem := NMHDRSize + (A_PtrSize * 2) + 4
        
        ; Get the address of the string buffer holding text from first column
        textAddr := NumGet(L + offText, "Ptr")
        
        ; Get the row we are over and then extract the text from the other columns
        row := NumGet(L + offItem, "Int") + 1
        snippet := this.snippets[row]
        
        SendMessage(TTM_SETTITLEW := 0x421, 0, snippet.GetAddress("title"), this.snippetsLVTooltipHwnd)
        StrPut(StrReplace(snippet.content, "`t", "    "), textAddr, "UTF-16")
    }
    
    snippetGui(modify := false) {
        if (modify) {
            if (!this.toModify := this.snippetsLV.GetNext()) {
                MsgBox("Please select a snippet to modify first", "No snippet selected", "Icon!")
                return
            }
        }
        else {
            this.toModify := false
        }
        
        if (!this.addSnippetGui) {
            this.addSnippetGui := GuiCreate("+ToolWindow", "Add new snippet")
            this.addSnippetGui.SetFont("s9", "Tahoma")
            this.addSnippetGui.addText("", "Title")
            this.addSnippetGui.addEdit("w400 vTitle")
            this.addSnippetGui.addText("", "Content")
            ; addSnippetGui.addEdit("w200 h200 WantTab t8 -Wrap vContent", modify ? this.snippets[toModify].content : "")
            this.edit := new Scintilla(this.addSnippetGui, "w400 h400 Border vContent", , 0, 0)
            setupSciControl(this.edit)
            
            addButtonW := 100
            addButtonX := (this.edit.ctrl.pos.w + (this.addSnippetGui.MarginX * 2) - addButtonW) // 2
            addButton := this.addSnippetGui.addButton("x" addButtonX " w" addButtonW " Default vSubmit", "")
            addButton.OnEvent("Click", (ctrl) => this.handleClick(ctrl)) ;modify ? this.modifySnippet(toModify, ctrl) : this.addSnippet(ctrl))
        }
        
        this.addSnippetGui.control["title"].value := modify ? this.snippets[this.toModify].title : ""
        this.edit.SetText("", modify ? this.snippets[this.toModify].content : "", 1)
        this.addSnippetGui.control["submit"].text := !modify ? "Add" : "Save"
        this.addSnippetGui.Show()
    }
    
    loadSnippets() {
        for i, snippet in this.snippets {
            this.snippetsLV.Add("", snippet["title"])
        }
    }
    
    handleClick(ctrl) {
        if (!data := this.submitAddSnippetGui(ctrl.gui)) {
            return
        }
        
        if (data["submit"] = "Add") {
            this.addSnippet(ctrl)
        }
        else if (data["submit"] = "Save") {
            if (!this.toModify) {
                MsgBox("Please select a snippet to modify first", "No snippet selected", "Icon!")
                return
            }
            this.modifySnippet(ctrl, this.toModify)
        }
    }
    
    addSnippet(ctrl) {
        this.snippets.push({id: this.snippets.length() + 1, title: data["title"], content: data["content"]})
        this.snippetsLV.Add("", data["title"])
    }
    
    modifySnippet(ctrl, index) {
        this.snippetsLV.Modify(index, "", data["title"])
        this.snippets[index].title := data["title"]
        this.snippets[index].content := data["content"]
    }
    
    removeSnippet() {
        lastSelectedIndex := 0
        while (id := this.snippetsLV.GetNext()) {
            this.snippets.RemoveAt(id)
            this.snippetsLV.Delete(id)
            lastSelectedIndex := id
        }
        if (lastSelectedIndex) {
            this.snippetsLV.Modify(lastSelectedIndex, "Select")
        }
    }
    
    submitAddSnippetGui(gui) {
        data := gui.submit(false)
        if (data["title"] = "" || data["content"] = "") {
            MsgBox("Title and Content cannot be empty.", "Missing data!")
            return false
        }
        else {
            gui.Destroy()
            return data
        }
    }
}