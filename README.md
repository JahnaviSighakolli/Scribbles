Public PreviousFullValues As Object

---

'Additional snapshot for fill-handle operations
Set PreviousFullValues = CreateObject("Scripting.Dictionary")

For Each cell In ws.UsedRange.Cells
    PreviousFullValues(cell.Address(False, False)) = GetCellContent(cell)
Next cell

--

If PreviousValues.Exists(cellAddress) Then

    oldValue = PreviousValues(cellAddress)

ElseIf Not PreviousFullValues Is Nothing Then

    If PreviousFullValues.Exists(cellAddress) Then
        oldValue = PreviousFullValues(cellAddress)
    Else
        oldValue = "[BLANK]"
    End If

Else

    oldValue = "[BLANK]"

End If

---

If Not PreviousFullValues Is Nothing Then
    PreviousFullValues(cellAddress) = newValue
End If
