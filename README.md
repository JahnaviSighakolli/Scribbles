
Public Sub UndoLastChange()

    Dim record As clsUndoRecord
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim cell As Range
    Dim oldData As Variant
    Dim i As Long
    Dim nextRow As Long
    Dim currentValue As String
    Dim restoredValue As String
    Dim previousEvents As Boolean

    On Error GoTo ErrorHandler

    InitializeUndoManager

    If UndoHistory.Count = 0 Then Exit Sub

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    IsPerformingUndo = True

    'Get the latest Undo record
    Set record = UndoHistory(UndoHistory.Count)

    Set ws = ThisWorkbook.Worksheets(record.SheetName)
    Set logWs = GetOrCreateChangeLog()

    'Remove this record from the Undo history
    UndoHistory.Remove UndoHistory.Count

    'Restore the original contents
    For i = LBound(record.CellAddresses) _
        To UBound(record.CellAddresses)

        Set cell = ws.Range(record.CellAddresses(i))

        oldData = record.OldContents(i)

        currentValue = GetCellContent(cell)

        If oldData(0) = True Then
            cell.Formula = oldData(1)
        ElseIf IsEmpty(oldData(1)) Then
            cell.ClearContents
        Else
            cell.Value2 = oldData(1)
        End If

        restoredValue = GetCellContent(cell)

        'Record the reversal in the audit log
        nextRow = logWs.Cells(logWs.Rows.Count, "A") _
                    .End(xlUp).Row + 1

        With logWs
            .Cells(nextRow, 1).Value = Now
            .Cells(nextRow, 2).Value = Environ$("Username")
            .Cells(nextRow, 3).Value = ws.Name
            .Cells(nextRow, 4).Value = cell.Address(False, False)
            .Cells(nextRow, 5).Value = LogValue(currentValue)
            .Cells(nextRow, 6).Value = LogValue(restoredValue)
        End With

    Next i

CleanExit:

    IsPerformingUndo = False
    Application.EnableEvents = previousEvents

    'Register the next Undo operation
    If UndoHistory.Count > 0 Then
        Application.OnUndo _
            "Undo last logged change", _
            "'" & ThisWorkbook.Name & "'!UndoLastChange"
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Undo Error: " & Err.Description, vbExclamation

    Resume CleanExit

End Sub
