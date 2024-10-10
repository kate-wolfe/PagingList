Attribute VB_Name = "MainModule"
Option Explicit

Sub Main()

'Turn off screen animations to make things go faster.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'On Error Resume Next

'If you need to troubleshoot, it may be worth going through each procedure one by one in the order they are called below.
'To find a procedure in this module, go to the dropdown in the top right and choose it.

'Run the first procedure, which takes the pasted email and fills in the Complete sheet based on it.
Call GetComplete

'This procedure sorts the Complete sheet into the rest of the sheets.
Call Sort

'Split Complete list between red bins and grey bins
Call BinSplit

'Format sheets
Call Format

'Save Stats
Call Stats

End Sub

Sub GetComplete()

'Define variables. This allocates memory and lets us use shorter names for sheets and things.
Dim Email As Worksheet, Complete As Worksheet
Dim lastRow As Integer, emailRow As Integer, compRow As Integer
Dim resultCell As Object

Set Email = Sheets("Paste Email Here")
Set Complete = Sheets("Complete")


'Clear any formatting.
Email.UsedRange.ClearFormats


'Find the last row so we can loop only to that place, and not the entire sheet.
lastRow = Email.Range("A1").SpecialCells(xlCellTypeLastCell).Row


'emailRow is set to 1 initially to start at the first row of the sheet.
'compRow is set to 2 because there will be a header in row 1.
emailRow = 1
compRow = 2


'The next few lines of code loop through each row of the Email sheet and look for a barcode.
Do While emailRow <= lastRow

'"Find" function will return the cell where a barcode is found.
'That cell is stored in a variable called "resultCell"
Set resultCell = Email.Range(emailRow & ":" & emailRow).Find("31189")

 If resultCell Is Nothing Then 'If the "find" function does not find a barcode in the row then...
        emailRow = emailRow + 1 'go to the next row in the Email sheet.
    Else
        Complete.Cells(compRow, 4).Value2 = resultCell.Value2 'put barcode in column 4 (D)
        Complete.Cells(compRow, 1).Value2 = Right(resultCell.Value2, 4) 'put last 4 digits of barcode in column 1 (A)
        Complete.Cells(compRow, 2).Value2 = resultCell.Offset(-2).Value2 'look for call number two rows above barcode and put it in column 2 (B)
        Complete.Cells(compRow, 3).Value2 = resultCell.Offset(-1).Value2 'look for title one row above barcode and put in column 3 (C)
        Complete.Cells(compRow, 5).Value2 = resultCell.Offset(-3).Value2 'look for item location 3 rows above barcode and put in column 5 (E)
        Complete.Cells(compRow, 6).Value2 = resultCell.Offset(1).Value2 'look for pickup location one row below barcode and put in column 6 (F)
        emailRow = emailRow + 1 'go to the next row in the Email sheet.
        compRow = compRow + 1 'go to the next row in the Complete sheet to enter more data.
    End If
Loop

'At this point the loop through the Email sheet is done and the Complete sheet should have the complete paging list.

End Sub

Sub Sort()

'Define variables. This allocates memory and lets us use shorter names for sheets and things.
Dim Complete As Worksheet, NewBooks As Worksheet, Mezz As Worksheet, L1 As Worksheet, Stone As Worksheet, FloorTwo As Worksheet
Dim lastRow As Integer, compRow As Integer, newRow As Integer, mezzRow As Integer, lowerRow As Integer, stoneRow As Integer, twoRow As Integer
Dim location As String, callNum As String

Set Complete = Sheets("Complete")
Set NewBooks = Sheets("New")
Set Mezz = Sheets("Mezzanine")
Set L1 = Sheets("L1")
Set Stone = Sheets("Stone")
Set FloorTwo = Sheets("2nd Floor")


'Find the last row so we can loop only to that place, and not the entire sheet.
lastRow = Complete.Range("A1").SpecialCells(xlCellTypeLastCell).Row

compRow = 2
newRow = 2
mezzRow = 2
lowerRow = 2
stoneRow = 2
twoRow = 2

Do While compRow <= lastRow

location = Complete.Cells(compRow, 5).Value2
callNum = Complete.Cells(compRow, 2).Value2

Select Case True
  Case location = "New Books", _
       location = "Tech/Equipment", _
       location = "Adult" And callNum Like "MAGAZINE*"
    'put data in New Books sheet
    NewBooks.Range("C" & newRow & ":F" & newRow).Value2 = Complete.Range("A" & compRow & ":D" & compRow).Value2
    newRow = newRow + 1
  Case location = "Audiovisual", _
       location = "Large Print", _
       location = "Literacy"
    Mezz.Range("C" & mezzRow & ":F" & mezzRow).Value2 = Complete.Range("A" & compRow & ":D" & compRow).Value2
    mezzRow = mezzRow + 1
  Case location = "Paperback", _
       location = "Adult" And callNum Like "FICTION*", _
       location = "Adult" And callNum Like "GRAPHIC*", _
       location = "Adult" And callNum Like "MANGA*"
    L1.Range("C" & lowerRow & ":F" & lowerRow).Value2 = Complete.Range("A" & compRow & ":D" & compRow).Value2
    lowerRow = lowerRow + 1
  Case location = "Adult" And callNum Like "MYSTERY*", _
       location = "Adult" And callNum Like "SCI FIC*"
    Stone.Range("C" & stoneRow & ":F" & stoneRow).Value2 = Complete.Range("A" & compRow & ":D" & compRow).Value2
    stoneRow = stoneRow + 1
  Case location = "Adult"
    FloorTwo.Range("C" & twoRow & ":F" & twoRow).Value2 = Complete.Range("A" & compRow & ":D" & compRow).Value2
    twoRow = twoRow + 1
End Select

compRow = compRow + 1

Loop

With Complete
    .Cells(1, 1).Value2 = "Last4"
    .Cells(1, 2).Value2 = "Call Number"
    .Cells(1, 3).Value2 = "Title"
    .Cells(1, 4).Value2 = "Barcode"
    .Cells(1, 5).Value2 = "Location"
    .Cells(1, 6).Value2 = "Pickup Location"
    .Rows("1:1").Font.Bold = True
    .Columns("D:D").NumberFormat = "0"
    .Columns("C:C").NumberFormat = "0000"
End With


End Sub

Sub Format()

Dim ws As Worksheet
Dim wsName As Variant
Dim i As Long
Dim lastRow As Integer

wsName = Array("New", "Mezzanine", "L1", "Stone", "2nd Floor")

For i = LBound(wsName) To UBound(wsName)

Set ws = Worksheets(wsName(i))

lastRow = ws.Range("D1").SpecialCells(xlCellTypeLastCell).Row

    With ws.Range("D2:D" & lastRow)
        .Replace "FICTION", "FIC", xlPart
        .Replace "MYSTERY", "MYS", xlPart
        .Replace "SCI FIC", "SF", xlPart
        .Replace "SHORT STORIES", "SS", xlPart
        .Replace "CD CLASSICAL", "CD CLASS", xlPart
        .Replace "CD ROCK", "CD POP", xlPart
        .Replace "CD FOLK", "CD POP", xlPart
        .Replace "CD SNDTRK", "CD POP", xlPart
        .Replace "CD COUNTRY", "CD POP", xlPart
        .Replace "CD GENERAL", "CD POP", xlPart
        .Replace "CD POPULAR", "CD POP", xlPart
        .Replace "CDB MYS", "CDB FIC", xlPart
        .Replace "CDB SF", "CDB FIC", xlPart
        .Replace "[Great Courses]", "[G C]", xlPart
        .Replace "MP3", "CDB (MP3)", xlPart
        .Replace "ROMANCE", "ROM", xlPart
    End With

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
        .Cells(1, 1).Value2 = "Found"
        .Cells(1, 2).Value2 = "NOS"
        .Cells(1, 3).Value2 = "Last4"
        .Cells(1, 4).Value2 = "Call Number"
        .Cells(1, 5).Value2 = "Title"
        .Cells(1, 6).Value2 = "Barcode"
        .Rows("1:1").Font.Bold = True
        .Range("A1:F" & lastRow).RemoveDuplicates Columns:=Array(3, 6), Header:=xlYes
        .Columns("C:C").NumberFormat = "0000"
        .Columns("C:C").HorizontalAlignment = xlLeft
        .Columns("A:D").AutoFit
        .Columns("E:E").ColumnWidth = 35
        .Columns("F:F").NumberFormat = "0"
        .Columns("A:F").Font.Name = "Calibri"
        .Columns("A:F").Font.Size = 10
        .Columns("F:F").AutoFit
        .Columns("F:F").HorizontalAlignment = xlLeft
        .UsedRange.Borders.LineStyle = xlContinuous
        .UsedRange.Borders.Weight = xlThin
    End With
    
ws.Range("A1:F" & lastRow).Sort key1:=ws.Range("D1"), Order1:=xlAscending, Header:=xlYes

Next i


Dim ws2 As Worksheet
Dim wsName2 As Variant
Dim i2 As Long
Dim lastRow2 As Integer

wsName2 = Array("Local Holds", "MLN Holds")

For i2 = LBound(wsName2) To UBound(wsName2)

Set ws2 = Worksheets(wsName2(i2))

lastRow2 = ws2.Range("D1").SpecialCells(xlCellTypeLastCell).Row

    With ws2.Range("D2:D" & lastRow2)
        .Replace "FICTION", "FIC", xlPart
        .Replace "MYSTERY", "MYS", xlPart
        .Replace "SCI FIC", "SF", xlPart
        .Replace "SHORT STORIES", "SS", xlPart
        .Replace "CD CLASSICAL", "CD CLASS", xlPart
        .Replace "CD ROCK", "CD POP", xlPart
        .Replace "CD FOLK", "CD POP", xlPart
        .Replace "CD SNDTRK", "CD POP", xlPart
        .Replace "CD COUNTRY", "CD POP", xlPart
        .Replace "CD GENERAL", "CD POP", xlPart
        .Replace "CD POPULAR", "CD POP", xlPart
        .Replace "CDB MYS", "CDB FIC", xlPart
        .Replace "CDB SF", "CDB FIC", xlPart
        .Replace "[Great Courses]", "[G C]", xlPart
        .Replace "MP3", "CDB (MP3)", xlPart
        .Replace "ROMANCE", "ROM", xlPart
    End With
    
    With ws2.Range("H2:H" & lastRow2)
        .Replace "/Pickup", "", xlPart
    End With

    With ws2.PageSetup
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
    
    With ws2
        .Cells(1, 1).Value2 = "Found"
        .Cells(1, 2).Value2 = "NOS"
        .Cells(1, 3).Value2 = "Last4"
        .Cells(1, 4).Value2 = "Call Number"
        .Cells(1, 5).Value2 = "Title"
        .Cells(1, 6).Value2 = "Barcode"
        .Cells(1, 7).Value2 = "Location"
        .Cells(1, 8).Value2 = "Pickup"
        .Rows("1:1").Font.Bold = True
        .Range("A1:H" & lastRow2).RemoveDuplicates Columns:=Array(3, 6), Header:=xlYes
        .Columns("C:C").NumberFormat = "0000"
        .Columns("C:C").HorizontalAlignment = xlLeft
        .Columns("A:D").AutoFit
        .Columns("E:E").ColumnWidth = 35
        .Columns("F:F").NumberFormat = "0"
        .Columns("A:H").Font.Name = "Calibri"
        .Columns("A:H").Font.Size = 10
        .Columns("F:H").AutoFit
        .Columns("F:F").HorizontalAlignment = xlLeft
        .UsedRange.Borders.LineStyle = xlContinuous
        .UsedRange.Borders.Weight = xlThin
    End With
    
ws2.Range("A1:H" & lastRow2).Sort key1:=ws2.Range("G1"), Order1:=xlAscending, key2:=ws2.Range("D1"), Order2:=xlAscending, Header:=xlYes

Next i2


End Sub

Sub BinSplit()

Dim Complete As Worksheet, RedBins As Worksheet, GreyBins As Worksheet
Dim lastRow As Integer

Set Complete = Sheets("Complete")
Set RedBins = Sheets("Local Holds")
Set GreyBins = Sheets("MLN Holds")

lastRow = Complete.Range("A1").SpecialCells(xlCellTypeLastCell).Row

Complete.AutoFilterMode = False

'Red Bins first
Complete.Range("A:F").AutoFilter Field:=6, Criteria1:="=CAMBRIDGE*"
Complete.Range("A:F").AutoFilter Field:=5, Criteria1:=Array( _
        "Adult", "Audiovisual", "Large Print", "New Books", "Paperback", "Tech/Equipment"), _
        Operator:=xlFilterValues
        
Complete.Range("A1:F" & lastRow).SpecialCells(xlCellTypeVisible).Copy
RedBins.Range("C1").PasteSpecial xlPasteValuesAndNumberFormats

'Reset filter
Complete.AutoFilterMode = False

'Grey Bins next
Complete.Range("A:F").AutoFilter Field:=6, Criteria1:="<>*CAMBRIDGE*"
Complete.Range("A:F").AutoFilter Field:=5, Criteria1:=Array( _
        "Adult", "Audiovisual", "Large Print", "New Books", "Paperback", "Tech/Equipment"), _
        Operator:=xlFilterValues
        
Complete.Range("A1:F" & lastRow).SpecialCells(xlCellTypeVisible).Copy
GreyBins.Range("C1").PasteSpecial xlPasteValuesAndNumberFormats

Complete.AutoFilterMode = False

End Sub
