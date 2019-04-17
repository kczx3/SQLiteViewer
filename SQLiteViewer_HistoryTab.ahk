class SQLiteViewer_HistoryTab {
    groups := []
    groupCnt := 1
    
    sql := []
    
    __New(gui, height, width) {
        this.gui := gui
        this.LV := this.gui.addListView("w" width " h" height " Section Count100 -Multi -0x80 vHistoryLV", "Query|Rows")
        this.LV.ModifyCol(1, width - 100)
        this.LV.ModifyCol(2, "AutoHdr")
        SetExplorerTheme(this.LV.hwnd)
        LV_EX_EnableGroupView(this.LV.hwnd)

        this.hdrHwnd := LV_EX_GetHeader(this.LV.hwnd)
        WinSetStyle(-0x80, "ahk_id " this.hdrHwnd)
        
        this.LV.OnEvent("Click", (ctrl, row) => this.SetEditText(ctrl, row))
        
        this.edit := new Scintilla(gui, "wp-" gui.MarginX * 3 " hp ys Border vHistoryEdit", , 0, 0)
        setupSciControl(this.edit)
        this.edit.SetReadOnly(1)
    }
    
    setEditText(ctrl, row) {
        this.edit.SetReadOnly(0)
        row ? this.edit.SetText("", this.sql[row], 1) : this.edit.ClearAll()
        this.edit.SetReadOnly(1)
    }
    
    addRow(db, sql, rows) {
        if (!this.groups.HasKey(db)) {
            LV_EX_GroupInsert(this.LV.hwnd, this.groupCnt, db)
            LV_EX_GroupSetState(this.LV.hwnd, this.groupCnt, "Collapsible")
            this.groups[db] := this.groupCnt
            this.groupCnt++
        }
        
        ; replace any newlines and following white space with a single space for display in the ListView
        noNewLinesSql := RegExReplace(sql, "\R\s*", " ")
        rowNum := this.LV.Add("", noNewLinesSql, rows)
        
        this.sql[rowNum] := sql ; store for lookup later to display in readonly Scintilla control
        
        LV_EX_SetGroup(this.LV.hwnd, rowNum, this.groups[db])
    }
}