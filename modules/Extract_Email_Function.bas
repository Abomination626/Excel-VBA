Attribute VB_Name = "Extract_Email_Function"
Function ExtractEmails(inputRange As Range) As String
'******************************************
'THIS CODE CAN EXTRACT EMAIL FROM UNCLEAN
'DATA, USING REGEX.
'CAN BE DIRECTLY USED IN CELLS USING
'FUNCTION =ExtractEmails()
'******************************************
    Dim regex As Object
    Dim matches As Object
    Dim cell As Range
    Dim emailList As String
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
        .Global = True
    End With
    For Each cell In inputRange
        If regex.Test(cell.Value) Then
            Set matches = regex.Execute(cell.Value)
            For Each Match In matches
                emailList = emailList & Match.Value & ", "
                'emailList = emailList & Match.Value & vbNewLine
            Next Match
        End If
    Next cell
    If Len(emailList) > 0 Then
        emailList = Left(emailList, Len(emailList) - 2) ' Remove the trailing comma and space
    End If
    ExtractEmails = emailList
End Function

