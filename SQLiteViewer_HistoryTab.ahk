class SQLiteViewer_HistoryTab {
    groups := []
    groupCnt := 1
    
    __New(gui, height, width) {
        this.gui := gui
        this.LV := this.gui.addListView("w" width " h" height " Section Count100 -Multi LV0x4000 -0x80 vHistoryLV", "Query|Rows")
        this.LV.ModifyCol(1, width - 100)
        this.LV.ModifyCol(2, "AutoHdr")
        SetExplorerTheme(this.LV.hwnd)
        LV_EX_EnableGroupView(this.LV.hwnd)

        this.hdrHwnd := LV_EX_GetHeader(this.LV.hwnd)
        WinSetStyle(-0x80, "ahk_id " this.hdrHwnd)
        
        this.LV.OnEvent("Click", (ctrl, row) => this.SetEditText(row > 0 ? ctrl.GetText(row) : false))
        
        this.edit := new Scintilla(gui, "wp-" gui.MarginX * 3 " hp ys Border vHistoryEdit", , 0, 0)
        setupSciControl(this.edit)
        this.edit.SetReadOnly(1)
    }
    
    setEditText(text) {
        this.edit.SetReadOnly(0)
        text ? this.edit.SetText("", text, 1) : this.edit.ClearAll()
        this.edit.SetReadOnly(1)
    }
    
    addRow(db, sql, rows) {
        if (!this.groups.HasKey(db)) {
            LV_EX_GroupInsert(this.LV.hwnd, this.groupCnt, db)
            LV_EX_GroupSetState(this.LV.hwnd, this.groupCnt, "Collapsible")
            this.groups[db] := this.groupCnt
            this.groupCnt++
        }
        
        rowNum := this.LV.Insert(1, "", sql, rows)
        LV_EX_SetGroup(this.LV.hwnd, rowNum, this.groups[db])
    }
}