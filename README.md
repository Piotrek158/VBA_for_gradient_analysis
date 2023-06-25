# VBA_for_gradient_analysis
This VBA macro imports data from CSVs containing Stock prices history for individual stocks and performs gradient analysis and saves key values in newly created sheet

VBA code:
```VBA
Sub Main()
    Dim csvFolderPath As String
    Dim xlsxFolderPath As String
    Dim csvFile As String
    Dim xlsxFile As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    ' Set the folder paths for CSV and XLSX files
    csvFolderPath = "..."
    xlsxFolderPath = "..."
    
    ' Check if the CSV folder exists
    If Dir(csvFolderPath, vbDirectory) = "" Then
        MsgBox "CSV folder not found.", vbCritical
        Exit Sub
    End If
    
    ' Check if the XLSX folder exists, create it if not
    If Dir(xlsxFolderPath, vbDirectory) = "" Then
        MkDir xlsxFolderPath
    End If
    
    ' Loop through all CSV files in the folder
    csvFile = Dir(csvFolderPath & "*.csv")
    Do While csvFile <> ""
        ' Open CSV file and import data into a new worksheet
        Workbooks.OpenText Filename:=csvFolderPath & csvFile, DataType:=xlDelimited, Comma:=True
        
        ' Set a reference to the active worksheet
        Set ws = ActiveSheet
        
        ' Convert imported data into a table
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes)
        
        ' Rename the table to "data"
        tbl.Name = "data"
        
        ' Repair the date format in column A
        Dim dateColumn As Range
        Set dateColumn = tbl.ListColumns("Date").DataBodyRange
        
        For Each cell In dateColumn
            If InStr(cell.Value, "/") > 0 Then
                cell.Value = Format(DateValue(cell.Value), "dd.mm.yyyy")
            End If
        Next cell
        
        ' Generate the XLSX file name based on the original CSV file name
        xlsxFile = xlsxFolderPath & Left(csvFile, Len(csvFile) - 4) & ".xlsx"
        
        ' Save the entire table as an XLSX file
        tbl.Range.Copy
        Workbooks.Add
        ActiveSheet.Paste
        Application.CutCopyMode = False
        ActiveWorkbook.SaveAs Filename:=xlsxFile, FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
        
        ' Filter the table based on a specific date provided by the user
        Dim filterDate As Date
        filterDate = InputBox("Please enter the date to filter the table (yyyy-mm-dd):")
        
        ' Filter the table based on the date column
        With tbl.Range
            .AutoFilter Field:=1, Criteria1:="=" & Format(filterDate, "dd.mm.yyyy"), Operator:=xlAnd
        End With


        
        ' Create a new workbook for the filtered data
        Dim filteredWorkbook As Workbook
        Set filteredWorkbook = Workbooks.Add
        
        ' Copy the filtered data to the new workbook
        tbl.Range.SpecialCells(xlCellTypeVisible).Copy filteredWorkbook.Sheets(1).Range("A1")
        filteredWorkbook.Sheets(1).Name = "Sheet1"
        
        ' Generate the filtered XLSX file name based on the original CSV file name
        Dim filteredXlsxFile As String
        filteredXlsxFile = xlsxFolderPath & Left(csvFile, Len(csvFile) - 4) & "_filtered.xlsx"
        
        ' Save the filtered data as a new XLSX file
        filteredWorkbook.SaveAs Filename:=filteredXlsxFile, FileFormat:=xlOpenXMLWorkbook
        filteredWorkbook.Close SaveChanges:=False
        
        ' Close the CSV file
        ws.Parent.Close SaveChanges:=False
        
        Kill xlsxFile
        
        ' Move to the next CSV file
        csvFile = Dir
        
    Loop

    MsgBox "CSV files loaded into Excel tables and saved as XLSX files.", vbInformation
    
    
    Call analysis_call
End Sub


Sub analysis_call()
    Dim dictPath As String
    Dim filePath As String
    Dim wb As Workbook
    
    ' Set the directory path
    dictPath = "..."
    
    ' Loop through each XLSX file in the directory
    filePath = Dir(dictPath & "*.xlsx")
    Do While filePath <> ""
        ' Open the workbook
        Set wb = Workbooks.Open(dictPath & filePath)
        
        ' Add a new sheet named "Sheet2" before calling DataAnalysis2
        wb.Sheets.Add Before:=wb.Sheets(1)
        wb.Sheets(1).Name = "Sheet2"
        
        ' Call the desired macro in the workbook
        DataAnalysis2 wb
        
        ' Save and close the workbook
        wb.Close SaveChanges:=True
        
        ' Move to the next file
        filePath = Dir
    Loop
End Sub

Sub DataAnalysis2(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim ws1 As Worksheet
    Dim LastRow As Long
    Dim PeakRow As Long
    Dim DipRow As Long
    Dim PeakVal As Double
    Dim DipVal As Double
    Dim Checker As Integer
    Dim xfactor As Double
    Dim NineThirtyRow As Integer
    
    Checker = 0
    
    Set ws = wb.Sheets("Sheet1")
    Set ws1 = wb.Sheets("Sheet2")
    DipVal = ws.Cells(2, "J").Value
    LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    xfactor = (InputBox("Please enter the percent above which we consider peak/dip")) / 100
    ws1.Cells(2, "C").Value = ws.Cells(2, "H").Value
    
    For i = 2 To LastRow
        If ws.Cells(i, "B").Value = TimeValue("09:30:00 AM") Then
            ws1.Cells(3, "C").Value = ws.Cells(i - 1, "K").Value
            ws1.Cells(4, "C").Value = ws.Cells(i - 1, "T").Value
            ws1.Cells(8, "C").Value = ws.Cells(i + 1, "K").Value
            ws1.Cells(9, "C").Value = ws.Cells(i + 1, "T").Value
            NineThirtyRow = i
        End If
        
        If Checker Mod 2 = 0 Then
            If ws.Cells(i, "I").Value > ws.Cells(i + 1, "I").Value And ws.Cells(i, "I").Value > (DipVal * (1 + xfactor)) Then
                PeakRow = i
                PeakVal = ws.Cells(i, "I").Value
                ws.Cells(i, "W").Value = "Peak"
                ws.Cells(i, "X").Value = ws.Cells(i, "I").Value
                Checker = Checker + 1
            End If
        Else
            If ws.Cells(i, "J").Value < ws.Cells(i + 1, "J").Value And ws.Cells(i, "I").Value < (PeakVal * (1 - xfactor)) Then
                DipRow = i
                DipVal = ws.Cells(i, "J").Value
                ws.Cells(i, "W").Value = "Dip"
                ws.Cells(i, "X").Value = ws.Cells(i, "J").Value
                Checker = Checker + 1
            End If
        End If
    Next i
    
    ws1.Cells(5, "B").Value = "Minimum before 9:30"
    ws1.Cells(6, "B").Value = "Maximum before 9:30"
    ws1.Cells(10, "B").Value = "Minimum after 9:30"
    ws1.Cells(11, "B").Value = "Maximum after 9:30"
    ws1.Cells(13, "B").Value = "Cumulutative val"
    ws1.Cells(5, "C").Value = WorksheetFunction.Min(ws.Range("J2:J" & NineThirtyRow - 1))
    ws1.Cells(6, "C").Value = WorksheetFunction.Max(ws.Range("I2:I" & NineThirtyRow - 1))
    ws1.Cells(10, "C").Value = WorksheetFunction.Min(ws.Range("J" & NineThirtyRow + 1 & ":J" & LastRow))
    ws1.Cells(11, "C").Value = WorksheetFunction.Max(ws.Range("I" & NineThirtyRow + 1 & ":I" & LastRow))
    ws1.Cells(13, "C").Value = ws1.Cells(LastRow - 1, "T").Value

End Sub
```
