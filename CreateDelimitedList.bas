Option Explicit

'
'   CreateDelimitedList
'
'   Create a delimited list from the provided range of cells, using the provided
'   delimeter (default ",").
'
'   Blank cells are ignored.
'

Public Sub CreateDelimitedList()

    On Error GoTo err:

    Dim rngSource As Range
    Set rngSource = Application.Selection
    Set rngSource = Application.InputBox("Source:", , rngSource.Address, Type:=8)  'define the range
    
    Dim rngTarget As Range
    Set rngTarget = Application.Selection
    Set rngTarget = Application.InputBox("Output cell:", rngTarget.Address, Type:=8)   'define the range

    If WorksheetFunction.CountA(rngSource) > 0 Then

        rngTarget = rngTarget.Cells(0, 0)   'reduce the range to one cell
        rngTarget.Value = ""    'clear it
    
        Dim delimeter As String
        delimeter = Application.InputBox("Delimeter", Default:=",")   'define the delimeter
        
        Dim cell As Range
        Dim output As String
        
        For Each cell In rngSource  'loop through the source
        
            If Len(Trim(cell.Value)) > 0 Then   'if the cell isn't empty...
            
                If rngTarget.Value = "" Then
                    rngTarget.Value = cell.Value
                Else
                    rngTarget.Value = rngTarget.Value + "," + cell.Value
                End If
            
            End If
        
        Next
    
    Else
    
        MsgBox "Range empty - nothing to join"
    
    End If
    
    Exit Sub

'error handling
err:

    If err.Number = 424 Then    'object not set, inputbox has been cancelled
    
        MsgBox "Operation cancelled"
    
    Else
    
        MsgBox "error occurred: " + err.Description
    
    End If
    
    
End Sub

