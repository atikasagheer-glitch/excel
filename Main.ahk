#Requires AutoHotkey v2.0

excelFile := A_ScriptDir "\Database.xlsx"

xl := ComObject("Excel.Application")
xl.Visible := false

wb := xl.Workbooks.Open(excelFile)
ws := wb.Worksheets(1)

myGui := Gui()
myGui.Title := "Excel Database"
myGui.SetFont("s10")

LV := myGui.Add("ListView", "w700 h300 Grid", [
    "ID", "Name", "Category", "Quantity", "Price", "Total"
])

; Excel data → ListView
row := 2

while (ws.Cells(row, 1).Value != "") {
    id := ws.Cells(row, 1).Value
    name := ws.Cells(row, 2).Value
    category := ws.Cells(row, 3).Value
    quantity := ws.Cells(row, 4).Value
    price := ws.Cells(row, 5).Value
    total := quantity * price

    LV.Add("", id, name, category, quantity, price, total)

    row++
}

; Add button
addBtn := myGui.Add("Button", "x10 y320 w100", "Add")
addBtn.OnEvent("Click", AddRecord)

myGui.OnEvent("Close", (*) => ExitApp())
myGui.Show()

AddRecord(*) {
    global ws, wb, LV

    ib := InputBox("Enter Name:", "Add Record")
    if (ib.Result != "OK")
        return
    name := ib.Value

    ib := InputBox("Enter Category:", "Add Record")
    if (ib.Result != "OK")
        return
    category := ib.Value

    ib := InputBox("Enter Quantity:", "Add Record")
    if (ib.Result != "OK")
        return
    quantity := ib.Value

    ib := InputBox("Enter Price:", "Add Record")
    if (ib.Result != "OK")
        return
    price := ib.Value

    ; Find next empty Excel row
    row := 2
    while (ws.Cells(row, 1).Value != "")
        row++

    ; Generate ID
    id := row - 1

    ; Calculate Total
    total := quantity * price

    ; Write data to Excel
    ws.Cells(row, 1).Value := id
    ws.Cells(row, 2).Value := name
    ws.Cells(row, 3).Value := category
    ws.Cells(row, 4).Value := quantity
    ws.Cells(row, 5).Value := price
    ws.Cells(row, 6).Value := total

    ; Save Excel
    wb.Save()

    ; Add same record to ListView
    LV.Add("", id, name, category, quantity, price, total)

    MsgBox("Record added and Excel updated!")
}