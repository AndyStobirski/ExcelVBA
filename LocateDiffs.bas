Option Explicit

'
'   LocateDifference
'
'   Compare the two ranges and mark items in one that don't occur in the other
'
'   Cells are marked with a background colour

Public Sub LocateDifference()

    Dim rngCompare1 As Range
    Set rngCompare1 = Application.Selection
    Set rngCompare1 = Application.InputBox("Compare Source1:", , rngCompare1.Address, Type:=8)  'define the range
    
    Dim rngCompare2 As Range
    Set rngCompare2 = Application.Selection
    Set rngCompare2 = Application.InputBox("Compare Source2:", , rngCompare2.Address, Type:=8)  'define the range
    
    Dim diffsFound As Boolean
    diffsFound = False
           
    If compareRange(rngCompare1, rngCompare2, 65535) Then diffsFound = True
    
    If compareRange(rngCompare2, rngCompare1, 15773696) Then diffsFound = True
    
    If diffsFound Then
    
        MsgBox ("Diffs found and marked")
    
    Else
    
        MsgBox ("No diffs found")
    
    End If
    

End Sub

Function compareRange(rng1 As Range, rng2 As Range, markColour As Long) As Boolean

    Dim src, trg As Range
    Dim matchfound As Boolean
    Dim ctr As Integer
    ctr = 0
    matchfound = False

    For Each src In rng1
    
        matchfound = False
    
        'cell is not empty
        If Len(Trim(rng1(src.Row - rng1.Row + 1, src.Column - rng1.Column + 1).Value)) > 0 Then
        
            For Each trg In rng2
                    
                'cell is not empty
                If trg.Value = src.Value Then
                
                    matchfound = True
                    
                    Exit For
                
                End If
                
            Next
            
            If matchfound = False Then
                src.Interior.Color = markColour
                ctr = ctr + 1
            End If
                    
        End If
        
    Next
    
    If ctr > 0 Then
        compareRange = True
    Else
        compareRange = False
    End If
    
    

End Function

