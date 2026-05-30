# VBA Documentation for Inventory Dashboard

## Overview

This document explains the VBA macros and automation features used in the Inventory Management Dashboard.

## Module: InventoryAutomation.bas

### Main Procedures

#### 1. RefreshInventoryData

**Purpose**: Refreshes all data connections and recalculates the workbook.

```vba
Sub RefreshInventoryData()
    ' Refresh all data connections
    ThisWorkbook.RefreshAll
    
    ' Recalculate all formulas
    Application.CalculateFull
    
    ' Update timestamp
    Range("LastUpdated").Value = Now()
    
    MsgBox "Inventory data refreshed successfully!", vbInformation
End Sub
```

**Usage**: Click the "Refresh Data" button on the dashboard.

---

#### 2. GenerateReorderReport

**Purpose**: Creates a filtered report of items below reorder point and exports to PDF.

```vba
Sub GenerateReorderReport()
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim rng As Range
    Dim pdfPath As String
    
    Set ws = ThisWorkbook.Sheets("Inventory Data")
    Set wsReport = ThisWorkbook.Sheets("Reorder Report")
    
    ' Clear previous report
    wsReport.Cells.Clear
    
    ' Apply filter for items below reorder point
    ws.ListObjects("InventoryTable").Range.AutoFilter _
        Field:=5, Criteria1:="<=" & ws.Range("ReorderPointCell").Value
    
    ' Copy filtered data to report sheet
    ws.ListObjects("InventoryTable").SpecialCells(xlCellTypeVisible).Copy _
        Destination:=wsReport.Range("A1")
    
    ' Remove filter
    ws.ListObjects("InventoryTable").AutoFilter.ShowAllData
    
    ' Export to PDF
    pdfPath = ThisWorkbook.Path & "\Reorder_Report_" & Format(Date, "yyyymmdd") & ".pdf"
    wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
    
    MsgBox "Report generated: " & pdfPath, vbInformation
End Sub
```

**Usage**: Click the "Generate Reorder Report" button.

---

#### 3. UpdateStockLevels

**Purpose**: Updates stock levels based on recent transactions.

```vba
Sub UpdateStockLevels()
    Dim wsInv As Worksheet
    Dim wsTrans As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sku As String
    Dim qty As Double
    
    Set wsInv = ThisWorkbook.Sheets("Inventory Data")
    Set wsTrans = ThisWorkbook.Sheets("Transactions")
    
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through transactions and update inventory
    For i = 2 To lastRow
        sku = wsTrans.Cells(i, "A").Value
        qty = wsTrans.Cells(i, "B").Value
        
        ' Find SKU in inventory and update quantity
        With wsInv.Columns(1)
            Set found = .Find(What:=sku, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                wsInv.Cells(found.Row, "C").Value = _
                    wsInv.Cells(found.Row, "C").Value + qty
            End If
        End With
    Next i
    
    MsgBox "Stock levels updated!", vbInformation
End Sub
```

**Usage**: Run after importing new transaction data.

---

#### 4. SendLowStockAlert

**Purpose**: Sends email alert for low stock items (requires Outlook).

```vba
Sub SendLowStockAlert()
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim ws As Worksheet
    Dim lowStockItems As String
    Dim i As Long
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)
    Set ws = ThisWorkbook.Sheets("Inventory Data")
    
    ' Build list of low stock items
    lowStockItems = ""
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If ws.Cells(i, "C").Value <= ws.Cells(i, "E").Value Then
            lowStockItems = lowStockItems & _
                ws.Cells(i, "B").Value & " - Current: " & _
                ws.Cells(i, "C").Value & vbNewLine
        End If
    Next i
    
    ' Create email
    With mailItem
        .To = "purchasing@company.com"
        .Subject = "Low Stock Alert - " & Date
        .Body = "The following items are below reorder point:" & vbNewLine & _
                vbNewLine & lowStockItems
        .Display  ' Use .Send to send automatically
    End With
    
    Set mailItem = Nothing
    Set outlookApp = Nothing
End Sub
```

**Usage**: Schedule to run daily or run manually.

---

#### 5. ArchiveOldData

**Purpose**: Moves old transactions to archive sheet.

```vba
Sub ArchiveOldData()
    Dim wsTrans As Worksheet
    Dim wsArchive As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cutoffDate As Date
    
    Set wsTrans = ThisWorkbook.Sheets("Transactions")
    Set wsArchive = ThisWorkbook.Sheets("Archive")
    
    cutoffDate = DateAdd("m", -6, Date)  ' 6 months ago
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    
    ' Move rows older than cutoff date
    For i = lastRow To 2 Step -1
        If wsTrans.Cells(i, "D").Value < cutoffDate Then
            wsTrans.Rows(i).Cut Destination:=wsArchive.Rows(wsArchive.Rows.Count).End(xlUp).Offset(1)
        End If
    Next i
    
    MsgBox "Old data archived successfully!", vbInformation
End Sub
```

**Usage**: Run monthly for database maintenance.

---

## UserForm: frmNewItem

**Purpose**: Form for adding new inventory items.

### Form Controls

- `txtSKU`: Text box for SKU number
- `txtDescription`: Text box for item description
- `txtUnitCost`: Text box for unit cost
- `txtReorderPoint`: Text box for reorder point
- `cmdSave`: Command button to save
- `cmdCancel`: Command button to cancel

### Code Behind Form

```vba
Private Sub cmdSave_Click()
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = ThisWorkbook.Sheets("Inventory Data")
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Cells(nextRow, "A").Value = Me.txtSKU.Value
    ws.Cells(nextRow, "B").Value = Me.txtDescription.Value
    ws.Cells(nextRow, "D").Value = Me.txtUnitCost.Value
    ws.Cells(nextRow, "E").Value = Me.txtReorderPoint.Value
    
    MsgBox "New item added successfully!", vbInformation
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
```

---

## Event Handlers

### Workbook_Open Event

Automatically runs when workbook opens.

```vba
Private Sub Workbook_Open()
    ' Display welcome message
    MsgBox "Welcome to Inventory Management Dashboard!" & vbNewLine & _
           "Last updated: " & Range("LastUpdated").Value, vbInformation
    
    ' Refresh data on open
    Call RefreshInventoryData
End Sub
```

### Worksheet_Change Event

Triggers when inventory data changes.

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if changed cell is in quantity column
    If Not Intersect(Target, Me.Range("C:C")) Is Nothing Then
        ' Update conditional formatting
        Call UpdateStockStatus
    End If
End Sub
```

---

## Error Handling

All procedures include basic error handling:

```vba
On Error GoTo ErrorHandler
' ... code ...
Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Next
```

---

## Security Notes

1. **Enable Macros**: Users must enable macros to use automation features
2. **Digital Signature**: Consider signing macros with a digital certificate
3. **Password Protection**: Protect VBA project with password if distributing
4. **Trust Center**: Add workbook location to Trusted Locations

---

[Back to Project](../README.md)
