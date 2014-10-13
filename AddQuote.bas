Sub AddQuote()
Dim myCell As Range

    For Each myCell In Selection
        If myCell.Value <> "" Then
            myCell.Value = Chr(39) & Chr(39) & myCell.Value & Chr(39)
            If WorksheetFunction.CountA(Range("B1")) = 0 Then
                 myCell.Copy
                 Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone
            Else
                Dim policyNum
                Range("B1") = Range("B1") & ", " & myCell.Value
            End If
        End If
    Next myCell
End Sub

