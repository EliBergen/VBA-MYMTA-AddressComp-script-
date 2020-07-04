# VBA-MYMTA-AddressComp-script-
# One-off code to parse an Excel spreadsheet from work that hides agents who haven't changed addresses

Sub HideMatchingAddresses()

Dim i As Integer
Dim cell As Range

i = 258

For Each cell In Range("d258:d4015")

    dcontents = LCase(Range("D" & i).Value)
    jcontents = LCase(Range("J" & i).Value)

    pos1 = InStr(1, dcontents, " ")
    pos2 = InStr(1, jcontents, " ")
    
    address1 = Left(dcontents, pos1 - 1)
    address2 = Left(jcontents, pos2 - 1)
    
    If address1 = address2 Then
    cell.EntireRow.Hidden = True
    End If
    
    i = i + 1
Next cell


End Sub
