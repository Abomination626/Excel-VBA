Sub WindingOptrOutput()
Dim START, ROW_COUNT As Integer
START_ROW = 12
ROW_COUNT = 0
'SELECT SHEET
For I = 1 To Sheets.Count - 3

Sheets(I).Activate
'GET THE LAST ROW DATA OF SHEET
ROW_COUNT = Application.WorksheetFunction.CountA(ActiveSheet.Range("C12:c100"))
ROW_SUMMARY = Application.WorksheetFunction.CountA(Sheets("SUMMARY").Range("B1:B1000"))
'COPY NAME AND PART NO DATA
'COPY NAME AND PART NO DATA
'COPY THE FOUR SEGMENT DATA
Range("C" & START_ROW & ":D" & (ROW_COUNT + START_ROW - 1) & "," & _
      "G" & START_ROW & ":R" & (ROW_COUNT + START_ROW - 1) & "," & _
      "T" & START_ROW & ":V" & (ROW_COUNT + START_ROW - 1)).Select
      
'COPY THE TOTAL DATA
Selection.Copy

'GO TO NEXT SHEET
Sheets("SUMMARY").Activate

Range("A" & ROW_SUMMARY + 1 & ":A" & (ROW_SUMMARY + 1 + ROW_COUNT)).Value = Sheets(I).Name

Sheets("SUMMARY").Range("B" & ROW_SUMMARY + 1).Select
ActiveCell.PasteSpecial xlPasteValues

Next I
End Sub
