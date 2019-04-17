#Include <ZeeGrid>
class SQLiteViewer_ResultsListView {
    emptyText := ""
    emptyTextCenter := true
    
    columnDropDown := 0
    
    __New(gui, height, width, RS := False) {
        this.gui := gui
        
        
        ; this.grid := new ZeeGrid(gui, "x+0 y+0 w" width " h" height " vResultsLV Hidden")
        ; this.grid.SetGridBGColor(7)
        this.LV := gui.addListView("x+0 y+0 w" width " h" height " Count1000 -E0x200 LV0x14000 vResultsLV")
        this.LV.OnNotify(-312, (ctrl, l) => this.handleFilter(ctrl, l))
        this.LV.OnNotify(LVN_GETEMPTYMARKUP := -187, (ctrl, l) => this.setEmptyText(ctrl, l))
        this.LV.OnNotify(LVN_LINKCLICK := -184, () => Type(this.OnLinkClick) = "Func" && this.OnLinkclick() && this.resetEmptyText("Now run a query", true))
        this.LV.OnNotify(LVN_COLUMNDROPDOWN := -164, (ctrl, l) => this.handleDropDownClick(ctrl, l))
        
        SetExplorerTheme(this.LV.hwnd)
        HideFocusBorder(this.LV.hwnd)
        
        this.resultsLVHeaderHwnd := LV_EX_GetHeader(this.LV.hwnd)
        WinSetStyle("^0x180", "ahk_id " this.resultsLVHeaderHwnd)
        
        headerImgList := IL_Create(1)
        IL_Add(headerImgList, "primary_key.ico")
        HD_EX_SetImageList(this.resultsLVHeaderHwnd, headerImgList)
        
        this.menu := MenuCreate()
        this.menu.Add("Copy Column", () => this.copyColumn())
    }
    
    handleDropDownClick(ctrl, l) {
        this.columnDropDown := NumGet(l, A_PtrSize * 3 + 4) + 1
        
        VarSetCapacity(rect, 16)
        SendMessage(HDM_GETITEMDROPDOWNRECT := 0x1219, this.columnDropDown - 1, &rect, this.resultsLVHeaderHwnd)
        
        
        this.menu.show(NumGet(rect, "UInt") + this.LV.pos.x, NumGet(rect, 12, "UInt") + this.LV.pos.y)
    }
    
    copyColumn() {
        RS := this.RS
        
        str := ""
        if (RS.HasRows) {
            While(RC := this.RS.Next(Row) >= 1) {
                str .= (A_Index = 1 ? "" : "`n") . row[this.columnDropDown]
            }
            Clipboard := str
        }
    }
    
    setEmptyText(ctrl, l) {
        ; default text
        text := this.emptyText || "<a>Connect a database and run a query to get started!</a>"
        align := this.emptyTextCenter ? 1 : 0
        
        ; see https://docs.microsoft.com/en-us/windows/desktop/controls/lvn-getemptymarkup
        
        ; if showing default text, then center it in ListView
        NumPut(this.emptyTextCenter, l, A_PtrSize * 3, "UInt")
        
        ; put our string in the buffer
        StrPut(text, l + A_PtrSize * 3 + 4, StrLen(text))
        return true
    }
    
    resetEmptyText(text, center := false) {
        this.emptyText := text
        this.emptyTextCenter := center
        SendMessage(LVM_RESETEMPTYTEXT := 0x1054, 0, 0, this.LV.hwnd)
    }
    
    redraw(flag) {
        this.LV.Opt((flag ? "+" : "-") "redraw")
    }
    
    clearRows(redraw := false) {
        if (redraw) {
            this.redraw(false)
        }
        
        this.LV.Delete()
        Loop(this.LV.getCount("Col")) {
            this.LV.DeleteCol(1)
        }
        
        if (redraw) {
            this.redraw(true)
        }
    }
    
    handleFilter(ctrl, l) {
        static HDM_GETITEM := 0x120B
        static NMHDR_Size := A_PtrSize * 3
        static HDItemSize := (4 * 6) + (A_PtrSize * 6)
        static typeOffset := 48
        static pvFilterOffset := 56
        
        hwnd := NumGet(l, 0, "Ptr")
        col := NumGet(l, NMHDR_Size, "Int")
        
        filters := []
        
        Loop(this.LV.GetCount("Column")) {
            ; String buffer to retrieve the filter text
            VarSetCapacity(filter, 64*2, 0)
            
            ; HDFILTERTEXT structure
            VarSetCapacity(HDTEXTFILTER, A_PtrSize + 4)
            NumPut(&filter, HDTEXTFILTER, 0, "UPtr") ; add pointer to string buffer variable
            NumPut(64*2, HDTEXTFILTER, 8, "Int") ; buffer size
            
            ; HDITEM struct
            VarSetCapacity(HDItem, HDItemSize)
            NumPut(0x100, HDItem, 0, "UInt") ; Set the Mask to HDI_FILTER := 0x100
            NumPut(0x0, HDItem, typeOffset, "UInt") ; Set the Type to HDFT_ISSTRING := 0x0
            NumPut(&HDTEXTFILTER, HDItem, pvFilterOffset, "Ptr") ; Add pointer to HDTEXTFILTER struct
            
            ; Send HDM_GETITEM and if successful, filter the ListView results
            if (SendMessage(HDM_GETITEM, A_Index - 1, &HDItem, , "ahk_id " hwnd)) {
                filterStr := StrGet(&filter, 64*2, "UTF-16")
                if (filterStr) {
                    filters.push({column: A_Index, filter: filterStr})
                }
            }
        }
        
        this.filterResults(filters)
    }
    
    filterResults(filters) {
        statusBar := this.gui.control["statusbar"]
        totalRows := 0
        filteredRows := 0
        
        for i, filter in filters {
            ; Can the filter be converted to a Number?  If so, do it
            isInteger := filter.filter is "Integer"
            isFloat := filter.filter is "Float"
            if (isInteger) {
                filter.filter := Integer(filter.filter)
            }
            else if (isFloat) {
                
            }
        }
        
        this.LV.Opt("-Redraw")
        this.LV.Delete()
        
        If (this.RS.HasRows) {
            While(RC := this.RS.Next(Row) >= 1) {
                totalRows++
                if (filters.length()) {
                    ; we use a number here because if there 
                    shouldAdd := 0
                    for i, filter in filters {
                        ; if the filter is a number, do a strict equality check, otherwise perform InStr comparison
                        if (Type(filter.filter) = "Integer" ? Row[filter.column] = filter.filter : InStr(Row[filter.column], filter.filter)) {
                            shouldAdd++
                        }
                    }
                    
                    ; the row must match on all filters
                    if (shouldAdd = filters.length()) {
                        filteredRows++, addRow()
                    }
                    else {
                        continue
                    }
                }
                else {
                    addRow()
                }
            }
        }
        
        this.RS.Reset() ; reset the results pointer
        
        this.LV.Opt("+Redraw")
        
        rowDisplay := (filters.length() ? filteredRows "/" : "") . totalRows " rows"
        statusBar.setText(rowDisplay, 3)
        
        this.resetEmptyText(!filteredRows ? "No rows matched your filters" : "")
        
        addRow() {
            rowNum := this.LV.Add()
            for i, val in Row {
                this.LV.Modify(rowNum, "Col" i, val)
            }
        }
    }
    
    setData(RS) {
        this.RS := RS
        
        this.redraw(False)
        this.clearRows()
        
        If (this.RS.HasNames) {
            this.RS.Next(Row)
            for i, column in this.RS.Columns {
                colType := "Text"
                if (this.RS.HasRows) {
                    if (Row[i] is "integer") {
                        colType := "Integer"
                    }
                    else if (Row[i] is "float") {
                        colType := "Float"
                    }
                }
                this.LV.InsertCol(i, "150 " colType, column.name)
                if (column.primaryKey) {
                    HD_EX_SetFormat(this.resultsLVHeaderHwnd, i, ["Image", "String", "SplitButton"])
                    HD_EX_SetImage(this.resultsLVHeaderHwnd, i, 1)
                }
                else {
                    HD_EX_SetFormat(this.resultsLVHeaderHwnd, i, ["String", "SplitButton"])
                }
            }
            this.RS.Reset()
        }

        If (this.RS.HasRows) {
            While(RC := this.RS.Next(Row) >= 1) {
                rowNum := this.LV.Add("", Row*)
            }
        }
        
        this.RS.Reset()
        
        this.resetEmptyText(rowNum ? "" : "Your query returned no results")

        this.redraw(True)
        return rowNum
    }
}