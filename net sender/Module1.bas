Attribute VB_Name = "Module1"

Public Function LoadListFromFile(ByRef SourceFile As String, _
     ByRef List As ListBox)
On Error GoTo ErrEvt
Dim TextLine As String, FN As Integer

List.Clear

FN = FreeFile
 Open SourceFile For Input As #FN ' Open file.
   Do While Not EOF(FN) ' Loop until end of file.
   Line Input #FN, TextLine ' Read line into variable.
   If TextLine <> LineToRem Then
    List.AddItem (TextLine)
   End If
Loop
Close #FN ' Close file.



ErrEvt:
Select Case Err.Number
   Case 51
      Err.Clear
   Case Else
End Select
Resume Next
End Function


Public Function LoadListFrolmFile(ByRef SourceFile As String, _
     ByRef List1 As ListBox)
On Error GoTo ErrEvt
Dim TextLine As String, FN As Integer

List.Clear

FN = FreeFile
 Open SourceFile For Input As #FN
   Do While Not EOF(FN)
   Line Input #FN, TextLine
   If TextLine <> LineToRem Then
    List.AddItem (TextLine)
   End If
Loop
Close #FN ' Close file.



ErrEvt:
Select Case Err.Number
   Case 51
      Err.Clear
   Case Else
End Select
Resume Next
End Function
