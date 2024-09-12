Attribute VB_Name = "Module1"
Sub WrapIFERROR()
'******************************************
'THIS CODE WRAPS THE CURRENT CELL FORMULA
'INTO A IFERROR FORMULA
'******************************************

    Dim cell As Range
    ' Loop through each cell in the selected range
    For Each cell In Selection
        ' Check if the cell contains a formula
        If cell.HasFormula Then
            ' Wrap the existing formula with IFERROR
            cell.Formula = "=IFERROR(" & Mid(cell.Formula, 2) & ","""")"
        End If
    Next cell
End Sub

