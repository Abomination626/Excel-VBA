Attribute VB_Name = "Custom_Formula"
Sub ApplyCustomFormula()
'******************************************
'THIS CODE WILL ASK USER FOR CUSTOM
'FORMULA TO PROCESS SELECTED CELLS
'******************************************
    Dim customFormula As String
    
    ' Prompt the user to enter the custom formula
    customFormula = InputBox("Enter the custom formula (e.g., /100+15):", "Custom Formula")
    
    For Each cell In Selection.Cells
    
    cell.Formula = "=" & cell.Value & customFormula
    
    Next cell
End Sub
