Attribute VB_Name = "mod_General"
'for ontop
'declares for ontop
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, Y, ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long

'load help file to section
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal nNewWord As Long) As Long

Global bAppLoading As Boolean
Global bWasAutoStart As Boolean

''''''''
Public Sub ShowHelp(Section As Integer)
Dim x
On Error GoTo EH
x = WinHelp(Form1.hwnd, App.Path & "\aspgen.hlp", cdlHelpContext, Section)
Exit Sub
EH:
MsgBox "Could not locate Help File.", vbCritical, "Charactor Creator"
Exit Sub
End Sub

Public Sub OnTop(Form As Object, Top As Boolean)
Dim Handle As Long
Dim Formx As Form
Set Formx = Form
Handle = Formx.hwnd
If Top = True Then SetWindowPos Handle, HWND_TOPMOST, 0, 0, 0, 0, _
TOPMOST_FLAGS

If Top = False Then SetWindowPos Handle, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub ShowError(Number As Long, Description As String, Title As String)
MsgBox Number & " - " & Description, vbCritical, Title
End Sub

Public Sub ShowInfo(Description As String, Title As String)
MsgBox Description, vbInformation, Title
End Sub

Public Function OpenTextFile(sFile As String) As String
'Reads an entire file into a string
On Error GoTo EH
Dim TMPTXT As String
Dim FinTxt As String
Dim iFile As Integer
iFile = FreeFile
Open sFile For Binary Access Read As #iFile
TMPTXT = Space$(LOF(iFile))
Get #iFile, , TMPTXT
Close #iFile
OpenTextFile = TMPTXT
Exit Function
EH:
OpenTextFile = ""
Exit Function
End Function

Public Function WriteTextFile(sFile As String, sData As String) As Boolean
On Error GoTo EH
Dim iFile As Integer
If CheckFile(sFile) = True Then
Kill sFile
End If
iFile = FreeFile

Open sFile For Binary Access Write As #iFile
Put #iFile, 1, sData
Close #iFile

WriteTextFile = True
Exit Function
EH:
WriteTextFile = False
Exit Function
End Function

Public Function ParsePath(ByVal TempPath As String, ReturnType As Integer)
'Parses a filename path
'Returns:
'Drive
'Directory
'Filename
'Extention

    Dim DriveLetter As String
    Dim DirPath As String
    Dim fname As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean

    If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 And ReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If

        DriveLetter = ""
        DirPath = ""
        fname = ""
        Extension = ""

        If Mid(TempPath, 2, 1) = ":" Then ' Find the drive letter.
            DriveLetter = Left(TempPath, 2)
            TempPath = Mid(TempPath, 3)
        End If

            PathLength = Len(TempPath)

            For Offset = PathLength To 1 Step -1 ' Find the next delimiter.
                Select Case Mid(TempPath, Offset, 1)
                 Case ".": ' This indicates either an extension or a . or a ..
                 ThisLength = Len(TempPath) - Offset

                 If ThisLength >= 1 Then ' Extension
                     Extension = Mid(TempPath, Offset, ThisLength + 1)
                 End If

                     TempPath = Left(TempPath, Offset - 1)
                     Case "\": ' This indicates a path delimiter.
                     ThisLength = Len(TempPath) - Offset

                     If ThisLength >= 1 Then ' Filename
                         fname = Mid(TempPath, Offset + 1, ThisLength)
                         TempPath = Left(TempPath, Offset)
                         FileNameFound = True
                         Exit For
                     End If

                         Case Else
                    End Select

                    Next Offset


                        If FileNameFound = False Then
                            fname = TempPath
                        Else
                            DirPath = TempPath
                        End If


                            If ReturnType = 0 Then
                                ParsePath = DriveLetter
                            ElseIf ReturnType = 1 Then
                                ParsePath = DirPath
                            ElseIf ReturnType = 2 Then
                                ParsePath = fname
                            ElseIf ReturnType = 3 Then
                                ParsePath = Extension
                            End If

End Function

Public Function CheckFile(sFile As String) As Boolean
'Does a file exist TRUE / FALSE
On Error Resume Next
If sFile = "" Then
CheckFile = False
Exit Function
End If
Dim Iret
Iret = Dir(sFile)
If Iret > "" Then
CheckFile = True
Else
If Iret = "" Then
CheckFile = False
End If
End If

End Function

Public Function GetLinEle(Origin As String, Sep1 As String, Sep2 As String) As String
'Parses a Line of text
On Error GoTo EH
Dim Bpos As Long
Dim Epos As Long
Bpos = InStr(1, Origin, Sep1, vbBinaryCompare)
If Bpos = 0 Then Exit Function
Epos = InStr(1, Origin, Sep2, vbBinaryCompare)
If Bpos = 0 Then Exit Function
Bpos = Bpos + Len(Sep1)
GetLinEle = Mid(Origin, Bpos, Epos - Bpos)
Exit Function
EH:
GetLinEle = ""
Exit Function
End Function


Public Sub LoadTree(ByVal tvTree As TreeView, ByVal sFileName As String)
    
    ' Function by Chetan Sarva (November 17,
    '     1999)
    ' Please include this comment if you use
    '     this code.
    Dim curNode As Node
    Dim sDelimiter As String
    Dim freef As Integer
    Dim buf As String
    Dim nodeparts As Variant
    sDelimiter = "" ' We want something extremely unique To delimit
    ' each of the pices of our treeview
    On Error Resume Next
    
    ' Get a free file and open our file for
    '     output
    freef = FreeFile()
    Open sFileName For Input As #freef
    


    Do


        DoEvents
            ' Read in the current line
            Line Input #freef, buf
            ' Split the line into pieces on our deli
            '     miter
            nodeparts = Split(buf, sDelimiter)
            
            ' See if it's a root or child node and a
            '     dd accordingly


            If nodeparts(3) = "parent" Then
                curNode = tvTree.Nodes.Add(, , nodeparts(1), nodeparts(0), CInt(nodeparts(4)))
                curNode.Tag = nodeparts(2)
                curNode.EnsureVisible
            Else
                curNode = tvTree.Nodes.Add(nodeparts(3), tvwChild, nodeparts(1), nodeparts(0), CInt(nodeparts(4)))
                curNode.Tag = nodeparts(2)
                curNode.EnsureVisible
            End If
            
        Loop Until EOF(freef)
        Close #freef
        
    End Sub


Public Sub SaveTree(ByVal tvTree As TreeView, ByVal sFileName As String)
    ' Function by Chetan Sarva (November 17,
    '     1999)
    ' Please include this comment if you use
    '     this code.
    Dim curNode As Node
    Dim sDelimiter As String
    Dim freef As Integer
    sDelimiter = "" ' We want something extremely unique To delimit
    ' each of the pices of our treeview
    On Error Resume Next
    
    ' Get a free file and open our file for
    '     output
    freef = FreeFile()
    Open sFileName For Output As #freef
    
    ' Loop through all the nodes and save al
    '     l the
    ' important information


    For Each curNode In tvTree.Nodes
        


        If curNode.FullPath = curNode.Text Then
            Print #freef, curNode.Text; sDelimiter; curNode.Key; sDelimiter; curNode.Tag; sDelimiter; "parent"; sDelimiter; curNode.Image
        Else
            Print #freef, curNode.Text; sDelimiter; curNode.Key; sDelimiter; curNode.Tag; sDelimiter; curNode.Parent.Key; sDelimiter; curNode.Image
        End If
        
    Next curNode
    Close #freef
    
End Sub


Public Sub SelectComboItem(sWhat As String, CMB As Object)
On Error GoTo EH
Dim i As Integer
For i = 0 To CMB.ListCount - 1
    If LCase(sWhat) = LCase(CMB.List(i)) Then
        CMB.ListIndex = i
        Exit For
    End If
Next i
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Selecting Combobox Item"
Exit Sub
End Sub

Public Function CheckTMPDir(sDir As String)
Dim Iret
Iret = Dir(sDir, vbDirectory)
If Iret = "" Then
MkDir sDir
End If
End Function
Public Sub RmTree(ByVal vDir As Variant)
'Removes a Directory structor
On Error Resume Next
Dim vFile As Variant
    ' Check if "\" was placed at end
    ' If So, Remove it
If Right(vDir, 1) = "\" Then
        vDir = Left(vDir, Len(vDir) - 1)
    End If
' Check if Directory is Valid
    ' If Not, Exit Sub
    vFile = Dir(vDir, vbDirectory)
If vFile = "" Then
        Exit Sub
    End If
' Search For First File
    vFile = Dir(vDir & "\", vbDirectory)
    ' Loop Until All Files and Directories
    ' Have been Deleted
Do Until vFile = ""


        If vFile = "." Or vFile = ".." Then
            vFile = Dir
        ElseIf (GetAttr(vDir & "\" & vFile) And _
            vbDirectory) = vbDirectory Then
            RmTree vDir & "\" & vFile
            vFile = Dir(vDir & "\", vbDirectory)
        Else
            Kill vDir & "\" & vFile
            vFile = Dir
        End If


    Loop


    ' Remove Top Most Directory
    RmDir vDir
End Sub

