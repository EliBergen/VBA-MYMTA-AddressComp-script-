# VBA-MYMTA-AddressComp-script-
# One-off code to parse an Excel spreadsheet from work that hides agents who haven't changed addresses

Sub DeleteMatchingAddresses()

Dim i As Integer
Dim numDeleted As Integer

i = 2
numDeleted = 0

Range("E2").Select

Do Until IsEmpty(ActiveCell)

    If ActiveCell.Value = "NULL" Then
        ActiveCell.Offset(1, 0).Select
        i = i + 1
        
    Else
        econtents = LCase(ActiveCell.Value)
        lcontents = LCase(Range("L" & i).Value)
        
        pos1 = InStr(1, econtents, " ")
        pos2 = InStr(1, lcontents, " ")
        
        address1 = Left(econtents, pos1 - 1)
        address2 = Left(lcontents, pos2 - 1)
    
        If address1 = address2 Then
            ActiveCell.EntireRow.Delete
            numDeleted = numDeleted + 1
            
        Else
            ActiveCell.Offset(1, 0).Select
            i = i + 1
    
    End If
    
End If

Loop

MsgBox "Rows deleted: " & numDeleted

End Sub

#Possible Addition

Sub HowManyHiddenRows()

Dim i As Integer

i = 0

Range("D2").Select

Do Until IsEmpty(ActiveCell)

    
    If ActiveCell.EntireRow.Hidden Then
        i = i + 1
        
    End If
    
    ActiveCell.Offset(1, 0).Select
    
Loop

MsgBox "Rows hidden: " & i

End Sub



