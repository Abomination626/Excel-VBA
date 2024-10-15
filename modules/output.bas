Public Sub CREATE_SUMMARY()
Dim START_ROW, ROW_COUNT As Integer
Dim SUMMARY_FLAG As Boolean

'INITIALIZE VARIABLE
START_ROW = 11
ROW_COUNT = 0
SUMMARY_FLAG = False


'SHOW ALL SHEETS
For i = 1 To Sheets.Count
    Sheets(i).Visible = True
    
    'CHECK IF SHEET(SUMMARY) EXIST
    If Sheets(i).Name = "SUMMARY_OUTPUT" Then
        SUMMARY_FLAG = True
    End If
Next i


'CREATE SHEET(SUMMARY) IF NOT EXIST
If SUMMARY_FLAG = False Then
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "SUMMARY_OUTPUT"
End If

For i = 1 To Sheets.Count
    If Mid(Sheets(i).Name, 1, 6) = "OUTPUT" Then
    Sheets(i).Activate
    
'Call GET_DATA
'================================================
'GET THE LAST ROW DATA OF SHEET
'ROW_COUNT = Application.WorksheetFunction.CountA(ActiveSheet.Range("C12:c100"))
'ROW_COUNT = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row - 12
ROW_COUNT = ActiveSheet.Range("C:C").SpecialCells(xlCellTypeLastCell).Row - 11

ROW_SUMMARY = Application.WorksheetFunction.CountA(Sheets("SUMMARY_OUTPUT").Range("B1:B1000"))

'COPY THE FOUR SEGMENT DATA
Range("C" & START_ROW & ":D" & (ROW_COUNT + START_ROW - 1) & "," & _
      "G" & START_ROW & ":R" & (ROW_COUNT + START_ROW - 1) & "," & _
      "T" & START_ROW & ":V" & (ROW_COUNT + START_ROW - 1)).Select
      
'COPY THE TOTAL DATA
Selection.Copy

'GO TO NEXT SHEET
Sheets("SUMMARY_OUTPUT").Activate

Range("A" & ROW_SUMMARY + 1 & ":A" & (ROW_SUMMARY + 1 + ROW_COUNT)).Value = Sheets(i).Name

Sheets("SUMMARY_OUTPUT").Range("B" & ROW_SUMMARY + 1).Select
ActiveCell.PasteSpecial xlPasteValues

'=================================================
    End If
Next i

Debug.Print "BREAKPOINT"

End Sub
