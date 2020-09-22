Attribute VB_Name = "mod_Files"
Global lColor As Long

Public Sub LoadActions()
Dim MyFile, MyPath, MyName
Form1.ListView1.ListItems.Clear
CheckTMPDir App.Path & "\temp\"
MyPath = App.Path & "\temp\"   ' Set the path.
MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
      If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
         Form1.ListView1.ListItems.Add , , MyName, , 4
      End If   ' it represents a directory.
   End If
   MyName = Dir   ' Get next entry.
Loop

End Sub

Public Sub LoadFrames(sName As String)
Dim MyFile, MyPath, MyName
Dim LST As ListItem
Form1.ListView2.ListItems.Clear

If CheckFile(App.Path & "\temp\" & sName & "\frames.dat") = True Then
    LoadFramesFromFile sName
    Exit Sub
End If

CheckTMPDir App.Path & "\temp\" & sName & "\"
MyPath = App.Path & "\temp\" & sName & "\"   ' Set the path.
MyName = Dir(MyPath, vbArchive)   ' Retrieve the first entry.
Do While MyName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If MyName <> "." And MyName <> ".." Then
      ' Use bitwise comparison to make sure MyName is a directory.
         Set LST = Form1.ListView2.ListItems.Add(, , "100", , 5)
         LST.SubItems(1) = MyName
   End If
   MyName = Dir   ' Get next entry.
Loop
'load the default times
End Sub

Public Sub LoadFramesFromFile(sName As String)
Dim TextLine
Dim iFile As Integer
Dim S() As String
Dim LST As ListItem
iFile = FreeFile
Open App.Path & "\temp\" & sName & "\frames.dat" For Input As #iFile   ' Open file.
Do While Not EOF(iFile)   ' Loop until end of file.
   Line Input #iFile, TextLine   ' Read line into variable.
    If TextLine <> "" Then
        S = Split(CStr(TextLine), ",")
        If UBound(S) > 0 Then
            Set LST = Form1.ListView2.ListItems.Add(, , S(0), , 5)
            LST.SubItems(1) = S(1)
        End If
    End If
Loop
Close #iFile   ' Close file.


End Sub

Public Sub WriteFramesFile(sName As String)
Dim sTMP As String
Dim bOK As Boolean
Dim i As Integer

For i = 1 To Form1.ListView2.ListItems.Count
    sTMP = sTMP & Form1.ListView2.ListItems(i).Text & "," & Form1.ListView2.ListItems(i).SubItems(1) & vbCrLf
Next i

bOK = WriteTextFile(App.Path & "\temp\" & sName & "\frames.dat", sTMP)
End Sub
Public Sub DoNew()
    CheckTMPDir App.Path & "\temp\"
    RmTree App.Path & "\temp\"
    CheckTMPDir App.Path & "\temp\"
    CheckTMPDir App.Path & "\temp\idle\"
    LoadActions
    Form1.ListView1.ListItems(1).Selected = True
    Form1.OldAction = "idle"
    LoadFrames Form1.ListView1.ListItems(1)
    lColor = vbBlack
End Sub

Public Sub DeleteAction(sName As String)
Dim i As Integer
If LCase(sName) = "idle" Then
        ShowInfo "You can not delete IDLE, this action must be present in all charactors.", "Warning"
    Exit Sub
End If
If DoesActionExist(sName) = False Then Exit Sub
For i = 1 To Form1.ListView1.ListItems.Count
    If LCase(Form1.ListView1.ListItems(i).Text) = LCase(sName) Then
        Form1.ListView1.ListItems.Remove i
    End If
Next i
RmTree App.Path & "\temp\" & sName & "\"
If Form1.ListView1.ListItems.Count > 0 Then
    Form1.ListView1.ListItems(1).Selected = True
    OldAction = Form1.ListView1.ListItems(1).Text
End If
End Sub


Public Sub AddAction(sName As String)
If DoesActionExist(sName) = True Then
    ShowInfo "Action Already Exists!", "User Error"
    Exit Sub
End If
CheckTMPDir App.Path & "\temp\" & sName & "\"
Form1.ListView1.ListItems.Add , , sName, , 4
End Sub

Public Function DoesActionExist(sName As String) As Boolean
Dim i As Integer
For i = 1 To Form1.ListView1.ListItems.Count
    If LCase(Form1.ListView1.ListItems(i).Text) = LCase(sName) Then
        DoesActionExist = True
        Exit Function
    End If
Next i
DoesActionExist = False
End Function

Public Sub AddFrames()
'On Error GoTo EH
If Form1.ListView1.SelectedItem Is Nothing Then
    ShowInfo "Please Select an Action!", "User Error"
    Exit Sub
End If

Dim Iret
Dim S() As String
Dim i As Integer
Dim LST As ListItem
Form1.commdlg.Filename = ""
Form1.commdlg.Filter = "BMP Files (*.bmp)|*.bmp"
Form1.commdlg.FilterIndex = 1
Form1.commdlg.Flags = cdlOFNAllowMultiselect
Form1.commdlg.DefaultExt = ".bmp"
Form1.commdlg.CancelError = True
Form1.commdlg.MaxFileSize = 32000
Form1.commdlg.InitDir = GetSetting(App.Title, "Frames", "LastPath", App.Path)
Form1.commdlg.ShowOpen
If Form1.commdlg.Filename = "" Then
    MsgBox "Please choose a filename!", vbInformation, "Please choose a file."
    Exit Sub
End If

'Form1.ActiveForm.Caption = Form1.commdlg.FileName
S = Split(Form1.commdlg.Filename, " ")
If UBound(S) = 0 Then
'they choose one file
    SaveSetting App.Title, "Frames", "LastPath", ParsePath(Form1.commdlg.Filename, 0) & ParsePath(Form1.commdlg.Filename, 1)
    Set LST = Form1.ListView2.ListItems.Add(, , "100", , 5)
    LST.SubItems(1) = ParsePath(Form1.commdlg.Filename, 2) & ParsePath(Form1.commdlg.Filename, 3)
    FileCopy Form1.commdlg.Filename, App.Path & "\temp\" & Form1.ListView1.SelectedItem.Text & "\" & ParsePath(Form1.commdlg.Filename, 2) & ParsePath(Form1.commdlg.Filename, 3)
Else
'they choose multiple files
SaveSetting App.Title, "Frames", "LastPath", S(0)
    For i = 1 To UBound(S)
        Set LST = Form1.ListView2.ListItems.Add(, , "100", , 5)
        LST.SubItems(1) = S(i)
        FileCopy S(0) & "\" & S(i), App.Path & "\temp\" & Form1.ListView1.SelectedItem.Text & "\" & S(i)
    Next i
Iret = MsgBox("Would you like to add the frames in reverse as well?", vbYesNo, "Add Reverse")
If Iret = vbYes Then
    For i = UBound(S) To 1 Step -1
        Set LST = Form1.ListView2.ListItems.Add(, , "100", , 5)
        LST.SubItems(1) = S(i)
    Next i
End If

End If

WriteFramesFile Form1.ListView1.SelectedItem.Text
Exit Sub
EH:
If Err <> cdlCancel Then
ShowError Err.Number, Err.Description, "Adding Frames"
End If
Exit Sub
End Sub

Public Sub AddActionSound()
On Error GoTo EH
If Form1.ListView1.SelectedItem Is Nothing Then
    ShowInfo "Please Select an Action!", "User Error"
    Exit Sub
End If

Dim Iret
Dim i As Integer
Dim LST As ListItem
Form1.commdlg.Filename = ""
Form1.commdlg.Filter = "WAV Files (*.wav)|*.wav"
Form1.commdlg.FilterIndex = 1
Form1.commdlg.Flags = cdlOFNExplorer + cdlOFNFileMustExist
Form1.commdlg.DefaultExt = ".wav"
Form1.commdlg.CancelError = True
Form1.commdlg.MaxFileSize = 32000
Form1.commdlg.InitDir = GetSetting(App.Title, "Frames", "LastPath", App.Path)
Form1.commdlg.ShowOpen
If Form1.commdlg.Filename = "" Then
    MsgBox "Please choose a filename!", vbInformation, "Please choose a file."
    Exit Sub
End If
Form1.ListView2.ListItems.Add 1, , "Sound", , 7
Form1.ListView2.ListItems(1).SubItems(1) = ParsePath(Form1.commdlg.Filename, 2) & ParsePath(Form1.commdlg.Filename, 3)

FileCopy Form1.commdlg.Filename, App.Path & "\temp\" & Form1.ListView1.SelectedItem.Text & "\" & ParsePath(Form1.commdlg.Filename, 2) & ParsePath(Form1.commdlg.Filename, 3)

Exit Sub
EH:
If Err <> cdlCancel Then
ShowError Err.Number, Err.Description, "Adding Action Sound"
End If
Exit Sub
End Sub

Public Sub DeleteFrame()
Dim i As Integer

For i = Form1.ListView2.ListItems.Count To 1 Step -1
If Form1.ListView2.ListItems(i).Selected = True Then
 DeleteFrameItem Form1.ListView2.ListItems(i)
End If
Next i
WriteFramesFile Form1.ListView1.SelectedItem.Text
End Sub

Public Sub DeleteFrameItem(LST As ListItem)
Dim i As Integer
Dim sFIle As String

If LST Is Nothing Then Exit Sub
sFIle = LST.SubItems(1)
Form1.ListView2.ListItems.Remove LST.Index

For i = 1 To Form1.ListView2.ListItems.Count
    If LCase(Form1.ListView2.ListItems(i).SubItems(1)) = LCase(sFIle) Then
    Exit Sub
    End If
Next i

'remove the frame file, no longer in use
Kill App.Path & "\temp\" & Form1.ListView1.SelectedItem.Text & "\" & sFIle
End Sub

Public Sub DoSave()
On Error GoTo EH
Dim BCZIP As New bcZipper
Dim sTMP As String
Dim i As Integer
Dim LST As ListItem
Dim sFIle As String
Dim sActionsList As String
Dim H As Integer
Set BCZIP = New bcZipper
Form1.commdlg.Filename = ""
Form1.commdlg.Filter = "Charactor Files (*.chr)|*.chr"
Form1.commdlg.FilterIndex = 1
Form1.commdlg.Flags = cdlOFNExplorer + cdlOFNFileMustExist
Form1.commdlg.DefaultExt = ".chr"
Form1.commdlg.CancelError = True
Form1.commdlg.MaxFileSize = 32000
Form1.commdlg.InitDir = App.Path
Form1.commdlg.ShowSave
sFIle = Form1.commdlg.Filename
WriteFramesFile Form1.ListView1.SelectedItem.Text
'Add the file header info
sActionsList = "My Agent Charactor File version " & App.major & App.minor & vbCrLf
sActionsList = sActionsList & "<chrdata>" & vbCrLf
sActionsList = sActionsList & "Color," & lColor & vbCrLf
sActionsList = sActionsList & "</chrdata>" & vbCrLf
sActionsList = sActionsList & "<chractions>" & vbCrLf
For i = 1 To Form1.ListView1.ListItems.Count
sActionsList = sActionsList & Form1.ListView1.ListItems(i).Text & vbCrLf
sTMP = sTMP & "<" & Form1.ListView1.ListItems(i).Text & ">" & vbCrLf
    If CheckFile(App.Path & "\temp\" & Form1.ListView1.ListItems(i).Text & "\frames.dat") = True Then
        sTMP = sTMP & OpenTextFile(App.Path & "\temp\" & Form1.ListView1.ListItems(i).Text & "\frames.dat")
    End If
sTMP = sTMP & "</" & Form1.ListView1.ListItems(i).Text & ">" & vbCrLf
Next i
sActionsList = sActionsList & "</chractions>" & vbCrLf

'write the info file
WriteTextFile App.Path & "\temp\info.dat", sActionsList & sTMP

'compress the files
If BCZIP.ZipaFolderEX(App.Path & "\temp\", App.Path & "\tcompress.zip") = True Then
    DoEvents
    If CheckFile(App.Path & "\tcompress.zip") = True Then
        FileCopy App.Path & "\tcompress.zip", sFIle
        Kill App.Path & "\tcompress.zip"
    End If
End If

Exit Sub
EH:
If Err <> cdlCancel Then
ShowError Err.Number, Err.Description, "Saving Charactor"
End If
Exit Sub
End Sub

Public Sub DoOpen()
Dim BCZIP As New bcZipper
Dim sTMP As String
Dim i As Integer
Dim LST As ListItem
Dim sFIle As String
Dim sActionsList As String
Dim sColor As String
Dim H As Integer
Dim S() As String
Dim sField As String
Set BCZIP = New bcZipper
Form1.commdlg.Filename = ""
Form1.commdlg.Filter = "Charactor Files (*.chr)|*.chr"
Form1.commdlg.FilterIndex = 1
Form1.commdlg.Flags = cdlOFNExplorer + cdlOFNFileMustExist
Form1.commdlg.DefaultExt = ".chr"
Form1.commdlg.CancelError = True
Form1.commdlg.MaxFileSize = 32000
Form1.commdlg.InitDir = App.Path
Form1.commdlg.ShowOpen
sFIle = Form1.commdlg.Filename

DoNew

If CheckFile(App.Path & "\tcompress.zip") = True Then
    Kill App.Path & "\tcompress.zip"
End If
    
FileCopy sFIle, App.Path & "\tcompress.zip"

If BCZIP.UnzipaFile(App.Path & "\tcompress.zip", App.Path & "\temp\") = True Then
    DoEvents
    LoadActions
    If Form1.ListView1.ListItems.Count > 0 Then
        LoadFrames Form1.ListView1.ListItems(1)
    End If
End If

If CheckFile(App.Path & "\temp\info.dat") = True Then
    sTMP = OpenTextFile(App.Path & "\temp\info.dat")
    S = Split(GetLinEle(sTMP, "<chrdata>", "</chrdata>"), vbCrLf)
    For i = LBound(S) To UBound(S)
        If InStr(1, LCase(S(i)), "color,") Then
            sColor = Mid(S(i), Len("color,") + 1, Len(S(i)))
            Exit For
        End If
    Next i
    If IsNumeric(sColor) = True Then
        lColor = CLng(sColor)
    Else
        lColor = vbBlack
    End If
End If

If CheckFile(App.Path & "\tcompress.zip") = True Then
    Kill App.Path & "\tcompress.zip"
End If


End Sub
