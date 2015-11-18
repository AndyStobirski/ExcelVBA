Option Explicit

'
'   LocateDupes
'
'   Locate duplicate entries in the selected range and mark each with
'   a number in brackets.
'
'   This may alter cells in the sheet being examined, so make a back up
'

Public Sub LocateDupes()

    Dim rngSrc As Range
    Set rngSrc = Application.Selection
    Set rngSrc = Application.InputBox("Source Ranges:", , rngSrc.Address, Type:=8)  'define the range
    Dim iMatch As Integer
    Dim cell, cellCheck As Range
    
    'a 2d array used to store whether a cell has been identified as a duplicate
    Dim bMatchArray() As Boolean
    ReDim bMatchArray(rngSrc.Rows.Count, rngSrc.Columns.Count) 'init to size of range

    Application.ScreenUpdating = False

    
    iMatch = 1
    
    For Each cell In rngSrc 'examine each cell
    
        'Check that current cell hasn't been marked as a duplication AND it isn't empty
        'Note: the offset of subtracting the origin of the source range (rngSrc.row and rngSrc.column)
        'from the current cell's row and column to locate in the bMatchArray
        If bMatchArray(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1) = False And _
            Len(Trim(rngSrc(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1).Value)) > 0 Then
    
            'Here, we check the above cell against all the others
            For Each cellCheck In rngSrc
            
                'check it hasn't been processed, it has the same value as the above cell but it isn't the same
                'address
                If bMatchArray(cellCheck.row - rngSrc.row + 1, cellCheck.Column - rngSrc.Column + 1) = False And _
                    cell.Value = cellCheck.Value And _
                    cell.Address <> cellCheck.Address Then
                                                                                
                    'Mark the matched cell in the array
                    If bMatchArray(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1) = False Then
                       bMatchArray(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1) = True
                    End If
                                                                                
                    'Mark the target cell with a number to indicate it's a dupe
                    rngSrc(cellCheck.row - rngSrc.row + 1, cellCheck.Column - rngSrc.Column + 1) = _
                        rngSrc(cellCheck.row - rngSrc.row + 1, cellCheck.Column - rngSrc.Column + 1) + " (" + CStr(iMatch) + ")"
                    
                    'mark it as being a dupe
                    bMatchArray(cellCheck.row - rngSrc.row + 1, cellCheck.Column - rngSrc.Column + 1) = True
                
                End If
                
            Next
            
        'We've found a match, so mark the original cell
        If bMatchArray(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1) Then
            
            rngSrc(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1).Value = _
                rngSrc(cell.row - rngSrc.row + 1, cell.Column - rngSrc.Column + 1).Value + " (" + CStr(iMatch) + ")"
            
            iMatch = iMatch + 1
        
        End If
        
        End If
                
    Next
    
    Application.ScreenUpdating = True
    
    If (iMatch > 1) Then
        MsgBox CStr(iMatch - 1) + " Matches found"
    End If
End Sub
