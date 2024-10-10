Attribute VB_Name = "SaveStats"
Sub Stats()

'Defining objects to save memory
'These are the variables we need to move the stats.
    Dim Paging As Workbook, Stats As Workbook
    Dim Complete As Worksheet, StatsSheet As Worksheet
    Dim StatsDate As Range, StatsLast As Range
    Dim Total As Integer
    Dim FolderPath As String
 
'Setting Ranges to shorten code
    Set Paging = ThisWorkbook
    Set Complete = ThisWorkbook.Sheets("Complete")

'This finds the last row in Complete with content and subtracts 1 from it to get the total number of items on the list.
'You subtract 1 due to the header taking up a row.
    Total = Complete.Range("A1").SpecialCells(xlCellTypeLastCell).Row - 1
    
    
    FolderPath = Application.ActiveWorkbook.Path & "\Paging Stats.xlsm"
    Set Stats = Workbooks.Open(FolderPath)
    
    Set StatsSheet = Stats.Sheets("Stats")
    
'Find last column in Stats to enter the information.
    Set StatsDate = StatsSheet.Cells(StatsSheet.Rows.Count, "A").End(xlUp).Offset(1, 0)
    Set StatsLast = StatsSheet.Cells(StatsSheet.Rows.Count, "B").End(xlUp).Offset(1, 0)

'Add info to Stats.
StatsDate.Value2 = Date 'Today's date goes in column A
StatsLast.Value2 = Total 'The length of the Paging List goes in column B
    
Stats.Save
Stats.Close False

Complete.Activate
  
End Sub
