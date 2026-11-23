Debug.Print "NORMAL CHANGE PATH"
Debug.Print "Sheet: " & Sh.Name
Debug.Print "Target: " & Target.Address
Debug.Print "PreviousValues count: " & PreviousValues.Count
Debug.Print "PreviousFullValues count: " & PreviousFullValues.Count

---


Debug.Print "INSERTION PATH: " & insertionType

---

Debug.Print "DELETION PATH: " & deletionType

---

Debug.Print "Cell: " & cellAddress
Debug.Print "In PreviousValues: " & PreviousValues.Exists(cellAddress)
Debug.Print "In PreviousFullValues: " & PreviousFullValues.Exists(cellAddress)

If PreviousFullValues.Exists(cellAddress) Then
    Debug.Print "Original value: " & PreviousFullValues(cellAddress)
End If

Debug.Print "Current value: " & newValue
