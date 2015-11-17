'
'   ConvertColumnsToColumn
'
'   17/11/2015
'
'   Microsoft Excel 2007 VBA Converts the selected range of multiple columns into one column,
'   offering the options to exclude blank cells and duplicated cells
'

Option Explicit

Dim iPasteIndex As Integer
Dim iIncludeDupes As Integer
Dim iIncludeBlankCells As Integer
Dim rngTrg As Range

'
'   Start - run this
'
Sub ConvertColumnsToColumn()

    iPasteIndex = 1
    
    Dim rngSrc As Range
    Dim iRow, iCol As Integer
    Dim sCellVal As String
        
    Set rngSrc = Application.Selection
    Set rngSrc = Application.InputBox("Source Ranges:", , rngSrc.Address, Type:=8)  'define the range
    Set rngTrg = Application.InputBox("Convert to (single cell):", , Type:=8)   'define the output
    iIncludeDupes = MsgBox("Include duplicate cells", vbYesNo)  'Include duplicates
    iIncludeBlankCells = MsgBox("Include blank cells", vbYesNo) 'Include blanks
    

    Application.ScreenUpdating = False
    
    'examine the source range a column at a time, working from top to bottom
    For iCol = 1 To rngSrc.Columns.Count
    
        For iRow = 1 To rngSrc.Rows.Count
            
            sCellVal = rngSrc.Cells(iRow, iCol).Value 'cell value
            
            Call Add(sCellVal)
            
        Next
        
    Next
        
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub

'
'   Perform the cell copy
'
Private Sub Add(pCellValue As String)

    'cell value blank and no blanks allows
    If Len(Trim(pCellValue)) = 0 And iIncludeBlankCells = vbNo Then Exit Sub    'bail

    Dim iTargetIndex As Integer
    
    'check  for dupes
    If iIncludeDupes = vbNo Then
    
        For iTargetIndex = 1 To iPasteIndex - 1
        
            If rngTrg.Offset(iTargetIndex, 0).Value = pCellValue Then Exit Sub  'bail
        
        Next
    
    End If
    
    'add the value
    rngTrg.Offset(iPasteIndex, 0).Value = pCellValue
    iPasteIndex = iPasteIndex + 1

End Sub
