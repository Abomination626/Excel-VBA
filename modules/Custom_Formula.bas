Attribute VB_Name = "Custom_Formula"
Sub ApplyCustomFormula()
'******************************************
'THIS CODE WILL ASK USER FOR CUSTOM
'FORMULA TO PROCESS SELECTED CELLS
'******************************************
    Dim customFormula As String
    Dim selectedRange As Range
    
    ' Prompt the user to enter the custom formula
    customFormula = InputBox("Enter the custom formula (e.g., /100+15):", "Custom Formula")
    
    ' Get the selected range
    Set selectedRange = Application.Selection
    
    ' Apply the custom formula to the selected cells
    selectedRange.Formula = "=" & selectedRange.Address & customFormula
End Sub
