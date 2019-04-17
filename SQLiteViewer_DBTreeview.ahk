class SQLiteViewer_DBTreeview {
    static tvStyles := 0x4 | 0x0020 | 0x400
    dbs := {}
    connectCallbacks := []
    disconnectCallbacks := []
    newWorksheetCallbacks := []
    
    __New(gui, height, width) {
        this.gui := gui
        
        static toolbarCommands := {
            1: () => this.promptForDb(),
            2: () => this.removeDb(),
            3: () => this.addNewWorksheet()
        }
        
        connectedDbsText := gui.addText("w200 Section", "Connected Databases")

        dbViewerImgList := IL_Create(2)
        IL_Add(dbViewerImgList, "add_database.ico")
        IL_Add(dbViewerImgList, "remove_database.ico")
        IL_Add(dbViewerImgList, "new_sqlite_worksheet.ico")

        ; save Y margin and set to 0
        prevMarginY := gui.MarginY
        gui.marginY := 0

        this.dbViewerToolbar := new Toolbar(gui, "h24 w200 Menu ToolTips Transparent")
        this.dbViewerToolbar.OnEvent("Click", (tb, id) => toolbarCommands.HasKey(id) && toolbarCommands[id](id))
        this.dbViewerToolbar.SetImageList(dbViewerImgList)
        this.dbViewerToolbar.ExStyle := 0x80 ; double buffer
        this.dbViewerToolbar.SetMaxTextRows()
        this.dbViewerToolbar.Add(, "Add Database", 0, , , , 1) ; Add Database Button
        this.dbViewerToolbar.Add(, "Remove Database", 1, "Disabled", , , 2) ; Remove Database Button
        this.dbViewerToolbar.Add()
        this.dbViewerToolbar.Add(, "New SQLite Worksheet", 2, "Disabled", , , 3) ; New SQLite Worksheet Button
        this.dbViewerToolbar.SetButtonSize(24, 24)
        this.dbViewerToolbar.AutoSize()
        
        this.createImageList()
        
        this.TV := this.gui.addTreeview("h" height - connectedDbsText.pos.h - this.dbViewerToolbar.ctrl.pos.h - 1 " w" width " vDBtreeview 0x200")
        SendMessage(0x112C, this.tvStyles, this.tvStyles, , "ahk_id " this.TV.hwnd)
        SetExplorerTheme(this.TV.hwnd)
        this.TV.SetImageList(this.TVImageList)
        this.TV.OnEvent("Click", (ctrl, id) => id > 0 && this.TV.Modify(id, "Select"))
        this.TV.OnEvent("DoubleClick", (ctrl, id) => this.loadColumns(id))
        this.TV.OnEvent("ItemSelect", (ctrl, id) => this.handleTVChange(ctrl, id))
        this.TV.OnNotify(TVN_BEGINDRAGW := -456, (ctrl, lParam) => this.handleDrag(ctrl, lParam))
        
        ; reset the Y margin
        gui.marginY := prevMarginY
    }
    
    handleDrag(ctrl, l) {
        static NMHDR_Size := A_PtrSize * 3
        static TVITEMSize := 48
        queryEdit := this.gui.control["query"]
        queryTabs.queryEdit.SetFocus(0)
        
        MouseGetPos(mouseX, mouseY)
        
        ; Create the TVHITTESTINFO struct...
        VarSetCapacity( tvht, 16, 0 )
        NumPut( mouseX, tvht, 0, "int" )
        NumPut( mouseY - this.TV.Pos.Y - this.gui.MarginY + 5, tvht, 4, "int" )
        
        item := SendMessage(TVM_HITTEST := 0x1111, 0, &tvht, this.TV.hwnd)
        itemText := this.TV.GetText(item)
        
        hImgList := SendMessage(TVM_CREATEDRAGIMAGE := 0x1112, 0, item, this.TV.hwnd)
        
        DllCall("ImageList_BeginDrag", "Ptr", hImgList, "Int", 0, "Int", 0, "Int", 0)
        DllCall("ImageList_DragEnter", "Ptr", this.gui.hwnd, "Int", mouseX, "Int", mouseY)
        
        While GetKeyState("LButton") {
            MouseGetPos(mouseX, mouseY, , hoverCtrl, 3)
            
            ; move the image with the mouse
            DllCall("ImageList_DragMove", "Int", mouseX, "Int", mouseY)
        }
        
        if (hoverCtrl = queryEdit.hwnd) {
            VarSetCapacity(point, 24, 0)
            DllCall("User32.dll\GetCursorPos", "Ptr", &point)
            DllCall("User32.dll\ScreenToClient", "Ptr", queryEdit.hwnd, "Ptr", &point)
            charPos := queryTabs.queryEdit.PositionFromPoint(NumGet(point), NumGet(point, 4)) ; queryTabs is global
            this.insertTextIntoQuery(this.gui.control["query"], itemText, charPos)
        }
        
        DllCall("ImageList_EndDrag")
        DllCall("ImageList_DragLeave", "Ptr", this.gui.hwnd)
    }
    
    insertTextIntoQuery(ctrl, text, pos) {
        length := StrLen(text)
        
        if (GetKeyState("Ctrl") && !(RegExMatch(Chr(queryTabs.queryEdit.GetCharAt(pos)), "\s") || RegExMatch(Chr(queryTabs.queryEdit.GetCharAt(pos - 1)), "\s"))) {
            start := queryTabs.queryEdit.WordStartPosition(pos)
            end := queryTabs.queryEdit.WordEndPosition(pos)
            queryTabs.queryEdit.DeleteRange(start, end - start)
            pos -= pos - start ; adjust the position by the difference of the original position and the Word start
        }
        
        queryTabs.queryEdit.InsertText(pos, text, 1)
        
        queryTabs.queryEdit.ctrl.focus()
        queryTabs.queryEdit.SetFocus(1)
        queryTabs.queryEdit.GoToPos(pos += length)
    }
    
    handleTVChange(ctrl, id) {
        if (id && !this.TV.GetParent(id)) {
            this.dbViewerToolbar.EnableButton(2, true)
            this.dbViewerToolbar.EnableButton(3, true)
        }
        else {
            if (this.dbViewerToolbar.isButtonEnabled(2)) {
                this.dbViewerToolbar.EnableButton(2, false)
                this.dbViewerToolbar.EnableButton(3, false)
            }
        }
    }
    
    createImageList() {
        this.TVImageList := IL_Create(1)
        this.databaseIcon := IL_Add(this.TVImageList, "database.ico")
        this.tableIcon := IL_Add(this.TVImageList, "database_table.ico")
        this.ColumnIcon := IL_Add(this.TVImageList, "table_column.ico")
        this.keyIcon := IL_Add(this.TVImageList, "primary_key.ico")
    }
    
    redraw(flag) {
        this.TV.Opt((flag ? "+" : "-") "Redraw")
    }
    
    regexp(DB, ArgC, vals) {
        regexNeedle := StrGet(DllCall("SQLite3.dll\sqlite3_value_text", "Ptr", NumGet(vals), "Cdecl Ptr"), "UTF-8")
        search := StrGet(DllCall("SQLite3.dll\sqlite3_value_text", "Ptr", NumGet(vals + A_PtrSize), "Cdecl Ptr"), "UTF-8")
        DllCall("SQLite3.dll\sqlite3_result_int", "Ptr", DB, "Int", RegexMatch(search, regexNeedle), "Cdecl") ; 0 = false, 1 = true
    }
    
    promptForDb() {
        if (!databaseFiles := FileSelect("M3", A_MyDocuments, "Select a database to add", "SQLite Database (*.db)")) {
            return false
        }
        
        files := StrSplit(databaseFiles, "`n")
        dir := files[1] . (SubStr(files[1], -1) = "\" ? "" : "\")
        files.RemoveAt(1)
        
        for i, file in files {
            DB := new SQLiteDB()
            
            if (!DB.OpenDB(dir . file)) {
                MsgBox("Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode, "SQLite Error")
                return
            }
            else {
                for i, cb in this.connectCallbacks {
                    cb.call(DB, file)
                }
                
                DB.createScalarFunction(regexp(db, argC, vals) => this.regexp(db, argC, vals), 2)
            }
            
            this.addDatabase(DB)
        }
        
        return true
    }
    
    addDatabase(db) {
        SplitPath(db._Path, fileName)
        id := this.TV.Add(fileName, , "Bold Icon" this.databaseIcon)
        this.dbs[id] := db
        this.loadTables(id)
        return id
    }
    
    removeDb() {
        selectedId := this.TV.GetSelection()
        if (!selectedID || this.TV.GetParent(selectedId)) {
            return
        }
        
        for i, cb in this.disconnectCallbacks {
            cb.Call(this.dbs[selectedId])
        }
        this.dbs[selectedId].CloseDB()
        this.dbs.Delete(selectedId)
        this.TV.Delete(selectedId)
        
        if (!this.TV.GetCount()) {
            if (this.dbViewerToolbar.isButtonEnabled(2)) {
                this.dbViewerToolbar.EnableButton(2, false)
                this.dbViewerToolbar.EnableButton(3, false)
            }
        }
    }
    
    addNewWorksheet() {
        selectedId := this.TV.GetSelection()
        if (selectedId && !this.TV.GetParent(selectedId)) {
            db := this.dbs[selectedId]
            SplitPath(db._Path, fileName)
            
            for i, cb in this.newWorksheetCallbacks {
                cb.Call(db, fileName)
            }
        }
    }
    
    loadTables(id) {
        if (!this.TV.GetChild(id)) {
            this.dbs[id].GetTable("SELECT name FROM sqlite_master WHERE type='table' and name not like 'sqlite_%';", RS)
            
            if (RS.HasRows) {
                this.redraw(false)
                while(RC := RS.Next(Row) >= 1) {
                    this.TV.Add(Row.1, id, "Sort Icon" this.tableIcon)
                }
                this.TV.Modify(id, "Expand")
                this.redraw(true)
            }
        }
    }
    
    loadColumns(id) {
        if (this.dbs.HasKey(dbId := this.TV.GetParent(id)) && !this.TV.GetChild(id)) {
            tableName := this.TV.GetText(id)
            
            this.dbs[dbId].GetTable("PRAGMA table_info(" tableName ")", RS)
            
            if (RS.HasRows) {
                this.redraw(false)
                while(RC := RS.Next(Row) >= 1) {
                    currCol := this.TV.Add(Row.2, id, "Icon" this.ColumnIcon)
                    this.TV.Add("Type - " Row.3, currCol, "Icon99")
                    this.TV.Add("Nullable? - " (Row.4 ? "YES" : "NO"), currCol, "Icon99")
                    this.TV.Add("Default Value - " Row.5, currCol, "Icon99")
                    this.TV.Add("Primary Key - " Row.6, currCol, Row.6 ? "Icon" this.keyIcon : "Icon99")
                }
                this.TV.Modify(id, "Expand")
                this.redraw(true)
            }
        }
    }
    
    OnConnect(cb) {
        this.connectCallbacks.push(cb)
    }
    
    OnNewWorksheet(cb) {
        this.newWorksheetCallbacks.push(cb)
    }
    
    OnDisconnect(cb) {
        this.disconnectCallbacks.push(cb)
    }
}