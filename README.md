
Public Sub LogCellChange(ByVal ws As Worksheet, _
                         ByVal Target As Range)

    Dim logSheet As Worksheet
    Dim cell As Range
    Dim oldValue As Variant
    Dim newValue As String
    Dim nextRow As Long
    Dim previousEvents As Boolean
    Dim cellAddress As String

    On Error GoTo ErrorHandler

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False

    'Check that a snapshot exists for this worksheet
    If PreviousSheet <> ws.CodeName Then GoTo CleanExit
    If PreviousValues Is Nothing Then GoTo CleanExit

    Set logSheet = GetOrCreateChangeLog()

    For Each cell In Target.Cells

        cellAddress = cell.Address(False, False)
        newValue = GetCellContent(cell)

        'Get the value captured before the edit
        If PreviousValues.Exists(cellAddress) Then
            oldValue = PreviousValues(cellAddress)
        Else
            oldValue = "[BLANK]"
        End If

        'Log only actual changes
        If Not ValuesAreSame(oldValue, newValue) Then

            nextRow = logSheet.Cells(logSheet.Rows.Count, "A") _
                              .End(xlUp).Row + 1

            With logSheet
                .Cells(nextRow, 1).Value = Now
                .Cells(nextRow, 2).Value = Environ$("Username")
                .Cells(nextRow, 3).Value = ws.Name
                .Cells(nextRow, 4).Value = cellAddress
                .Cells(nextRow, 5).Value = LogValue(oldValue)
                .Cells(nextRow, 6).Value = LogValue(newValue)
            End With

        End If

        'Refresh snapshots so another edit in the same selection
        'uses the latest values as its starting point.
        If PreviousValues.Exists(cellAddress) Then
            PreviousValues(cellAddress) = newValue
        End If

        If Not PreviousVerticalValues Is Nothing Then
            If PreviousVerticalValues.Exists(cellAddress) Then
                PreviousVerticalValues(cellAddress) = newValue
            End If
        End If

        If Not PreviousHorizontalValues Is Nothing Then
            If PreviousHorizontalValues.Exists(cellAddress) Then
                PreviousHorizontalValues(cellAddress) = newValue
            End If
        End If

    Next cell

CleanExit:
    Application.EnableEvents = previousEvents
    Exit Sub

ErrorHandler:
    MsgBox "Change Logger Error: " & Err.Description, _
           vbExclamation, "Change Logger"
    Resume CleanExit

End Sub
