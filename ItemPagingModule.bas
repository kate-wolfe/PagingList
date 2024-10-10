Attribute VB_Name = "ItemPagingModule"
Option Explicit

'This is the macro for setting up the Item Paging.
Sub ItemPaging()

'Turn off screen animations so this runs faster.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Define variables. This allocates memory and lets us use shorter names for sheets and things.
Dim ItemEmail As Worksheet, ItemList As Worksheet
Dim lastRow As Integer, itemRow As Integer, listRow As Integer
Dim resultCell As Object

Set ItemEmail = Sheets("Item List Email")
Set ItemList = Sheets("Item Paging List")


'Clear any formatting.
ItemEmail.UsedRange.ClearFormats


'Find the last row so we can loop only to that place, and not the entire sheet.
lastRow = ItemEmail.Range("A1").SpecialCells(xlCellTypeLastCell).Row


'emailRow is set to 1 initially to start at the first row of the sheet.
'compRow is set to 2 because there will be a header in row 1.
itemRow = 1
listRow = 2


'The next few lines of code loop through each row of the Email sheet and look for a barcode.
Do While itemRow <= lastRow

'"Find" function will return the cell where a barcode is found.
'That cell is stored in a variable called "resultCell"
Set resultCell = ItemEmail.Range(itemRow & ":" & itemRow).Find("31189")

 If resultCell Is Nothing Then 'If the "find" function does not find a barcode in the row then...
        itemRow = itemRow + 1 'go to the next row in the Item Email sheet.
    Else
        ItemList.Cells(listRow, 4).Value2 = resultCell.Value2 'put barcode in column 4 (D)
        ItemList.Cells(listRow, 1).Value2 = Right(resultCell.Value2, 4) 'put last 4 digits of barcode in column 1 (A)
        ItemList.Cells(listRow, 2).Value2 = resultCell.Offset(-1).Value2 'look for call number one row above barcode and put it in column 2 (B)
        ItemList.Cells(listRow, 3).Value2 = resultCell.Offset(-2).Value2 'look for title two rows above barcode and put in column 3 (C)
        ItemList.Cells(listRow, 5).Value2 = resultCell.Offset(2).Value2 'look for item location two rows below barcode and put in column 5 (E)
        ItemList.Cells(listRow, 6).Value2 = resultCell.Offset(3).Value2 'look for pickup location three rows below barcode and put in column 6 (F)
        itemRow = itemRow + 1 'go to the next row in the Item Email sheet.
        listRow = listRow + 1 'go to the next row in the Item List sheet to enter more data.
    End If
Loop


'This cleans up the cells.
With ItemList.UsedRange
.Replace "      BARCODE:  ", vbNullString, xlPart
.Replace "      CALL NO:  ", vbNullString, xlPart
.Replace "      TITLE:    ", vbNullString, xlPart
.Replace "      PICKUP AT:  ", vbNullString, xlPart
.Replace "      LOCATION:  CAMBRIDGE/", vbNullString, xlPart
End With


'Formatting
Dim ws As Worksheet
Set ws = Sheets("Item Paging List")

With ws.PageSetup
    .LeftHeader = "&A" 'Name of the sheet
    .CenterHeader = "&D &T" 'Date and Time the list was run
    .RightHeader = "&P of &N" 'Page number
    .LeftMargin = Application.InchesToPoints(0.5)
    .RightMargin = Application.InchesToPoints(0.5)
    .TopMargin = Application.InchesToPoints(0.6)
    .BottomMargin = Application.InchesToPoints(0.6)
    .HeaderMargin = Application.InchesToPoints(0.3)
    .FooterMargin = Application.InchesToPoints(0.3)
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = False
End With
    
With ws
    .Cells(1, 1).Value2 = "Last4"
    .Cells(1, 2).Value2 = "Call Number"
    .Cells(1, 3).Value2 = "Title"
    .Cells(1, 4).Value2 = "Barcode"
    .Cells(1, 5).Value2 = "Location"
    .Cells(1, 6).Value2 = "Pickup"
    .Rows("1:1").Font.Bold = True
    .Columns("A:F").Font.Name = "Calibri"
    .Columns("A:F").Font.Size = 10
    .Columns("A:A").NumberFormat = "0000"
    .Columns("A:A").HorizontalAlignment = xlLeft
    .Columns("A:B").AutoFit
    .Columns("C:C").ColumnWidth = 35
    .Columns("D:D").NumberFormat = "0"
    .Columns("D:F").AutoFit
    .Columns("D:D").HorizontalAlignment = xlLeft
    .UsedRange.Borders.LineStyle = xlContinuous
    .UsedRange.Borders.Weight = xlThin
End With
    
    
'Defining objects to save memory
'These are the variables we need to move the stats.
Dim Paging As Workbook, Stats As Workbook
Dim ItemSheet As Worksheet, StatsSheet As Worksheet
Dim StatsDate As Range, StatsLast As Range
Dim Total As Integer
Dim FolderPath As String
 
'Setting Ranges to shorten code
Set Paging = ThisWorkbook
Set ItemSheet = Paging.Sheets("Item Paging List")

'This finds the last row in Complete with content and subtracts 1 from it to get the total number of items on the list.
'You subtract 1 due to the header taking up a row.
Total = ItemSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row - 1

FolderPath = Application.ActiveWorkbook.Path & "\Paging Stats.xlsm"
Set Stats = Workbooks.Open(FolderPath)
    
Set StatsSheet = Stats.Sheets("Stats")
    
'Find last column in Stats to enter the information.
Set StatsDate = StatsSheet.Cells(StatsSheet.Rows.Count, "A").End(xlUp).Offset(1, 0)
Set StatsLast = StatsSheet.Cells(StatsSheet.Rows.Count, "B").End(xlUp).Offset(1, 0)

'Add info to Stats.
StatsDate.Value2 = Date 'Today's date goes in column A
StatsLast.Value2 = Total 'The length of the Paging List goes in column B
StatsLast.Offset(0, 1).Value2 = "Item Paging"
    
Stats.Save
Stats.Close False

ItemSheet.Activate

End Sub

