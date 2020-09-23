Attribute VB_Name = "Module1"
Public Const FileName = "C:\TMP.txt" 'can be left out or changed to whatever

Public Sub SaveTVToFile()

Open FileName For Output As #1

Print #1, "Root:" & Form1.TreeView1.Nodes(1).Key & "," & Form1.TreeView1.Nodes(1).Text

For i% = 2 To Form1.TreeView1.Nodes.Count
    Set nodX = Form1.TreeView1.Nodes(i%)
    Print #1, "Sub:" & nodX.Parent.Key & "," & nodX.Key & "," & nodX.Text
Next i%

Close #1

End Sub

Public Sub LoadTVFromFile()

Form1.TreeView1.Nodes.Clear

Open FileName For Input As #1
While Not EOF(1)
Line Input #1, Dummy

If Left$(Dummy, 5) = "Root:" Then
    RootNode$ = Mid$(Dummy, 6, Len(Dummy) - 5)
    RootKey$ = GetBefore(RootNode$)
    RootText$ = GetAfter(RootNode$)
    Set nodX = Form1.TreeView1.Nodes.Add(, , RootKey$, RootText$)
End If

If Left$(Dummy, 4) = "Sub:" Then
    SubNode$ = Mid$(Dummy, 5, Len(Dummy) - 4)
    SubRelation$ = GetBefore(SubNode$)
    TempSubKey$ = GetAfter(SubNode$)
    SubKey$ = GetBefore(TempSubKey$)
    SubText$ = GetAfter(TempSubKey$)
    Set nodX = Form1.TreeView1.Nodes.Add(SubRelation$, tvwChild, SubKey$, SubText$)
End If

Wend

For i% = 1 To Form1.TreeView1.Nodes.Count
    Form1.TreeView1.Nodes(i%).Expanded = True
Next i%

End Sub

Private Function GetBefore(Sentence As String) As String
Const Sign = ","
Dim Counter As Integer
Dim Before As String
Counter = 1
For Counter = 1 To Len(Sentence)
    If Mid(Sentence, Counter, 1) = Sign Then
        Exit For
    End If
Next Counter
If Counter <> Len(Sentence) Then
    Before = Left(Sentence, (Counter - 1))
Else
    Before = ""
End If
GetBefore = Before
End Function


Private Function GetAfter(Sentence As String) As String
Const Sign = ","
Dim Counter As Integer
Dim Rest As String
Counter = 1
For Counter = 1 To Len(Sentence)
    If Mid(Sentence, Counter, 1) = Sign Then
        Exit For
    End If
Next Counter
If Counter <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - Counter))
Else
    Rest = ""
End If
GetAfter = Rest
End Function


