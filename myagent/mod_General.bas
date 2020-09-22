Attribute VB_Name = "mod_General"
Global bCancel As Boolean
Global bLoaded As Boolean
Global hRgn As Long
Global bShowBallon As Boolean
Global fBalloon As frmBalloon
Global Lcolor As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Sub SetRegion(fRenderForm As Form, fDisplayForm As Form)
'Free the memory allocated by the previous Region
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from
'it, creating a new region

Lcolor = RGB(0, 0, 0) 'black is transparent color of frames
    hRgn = GetBitmapRegion(fRenderForm.Picture1.Picture, Lcolor)
'Set the Forms new Region
    fDisplayForm.Width = fRenderForm.Width
    fDisplayForm.Height = fRenderForm.Height
    SetWindowRgn fDisplayForm.hwnd, hRgn, True
    fDisplayForm.Picture = fRenderForm.Picture
End Sub

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
'wait for a certain amount of time  - in milliseconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait '* 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
    DoEvents
        If bCancel = True Then
            StopAllWavs
        Exit Do
        End If
        DoEvents
    Loop
End Function

Public Sub DoClose()
On Error GoTo EH
bCancel = True
bLoaded = False
DoEvents
Unload frmOffScreen
Unload frmRender
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Closing Charactor"
Exit Sub
End Sub

Public Function LoadCHRFile(sFile As String) As Boolean
'Load a *.CHR file (short for charactor)
'A CHR file is a zip file with the extention renamed to CHR
'We will just unzip the file
On Error GoTo EH
Dim bcZIP As bcZipper

Set bcZIP = New bcZipper

'create our temp dir
CheckTMPDir App.Path & "\temp\"
RmTree App.Path & "\temp\"
CheckTMPDir App.Path & "\temp\"
'copy the file to our temp dir
FileCopy sFile, App.Path & "\temp\chr.zip"
'unzip the file
LoadCHRFile = bcZIP.UnzipaFile(App.Path & "\temp\chr.zip", App.Path & "\temp\")

Exit Function
EH:
    ShowError Err.Number, Err.Description, "LoadCHRFile"
    LoadCHRFile = False
Exit Function
End Function

Public Sub DoAnimation(sName As String)
'Load a run an action animation
'Play sounds as needed
On Error GoTo EH
Dim bOK As Boolean
Dim sFrame() As String
Dim sActionData As String
Dim sFrameData() As String
Dim i As Integer

'Check to see if the action directory exists
If Dir(App.Path & "\temp\" & sName & "\", vbDirectory) = "" Then
    ShowInfo "Could not find:" & vbCrLf & App.Path & "\temp\" & sName & "\", "DoAnimation"
    Exit Sub
End If

'make sure the config file for the charactor is there
If CheckFile(App.Path & "\temp\info.dat") = False Then
    ShowInfo "Could not find:" & vbCrLf & App.Path & "\temp\info.dat", "DoAnimation"
    Exit Sub
End If

'Load the config file, then run the action
sActionData = GetLinEle(LCase(OpenTextFile(App.Path & "\temp\info.dat")), "<" & LCase(sName) & ">", "</" & LCase(sName) & ">")

'Make sure we actually have action data
If sActionData = "" Then
    ShowInfo "Action Table is Empty!", "DoAnimation"
    Exit Sub
End If

'get all the frames into an array
sFrame = Split(sActionData, vbCrLf)

'stop any playing sounds
StopAllWavs
'MsgBox "Starting" & vbCrLf & sActionData
'run the animation
For i = LBound(sFrame) To UBound(sFrame)
    If bCancel = True Then
        StopAllWavs
        Exit For
    End If
DoEvents
'ignore blank frame records
If sFrame(i) <> "" Then
    'get the current frame data
    sFrameData = Split(sFrame(i), ",")
    'MsgBox "Frame Data loaded" & vbCrLf & LBound(sFrameData) & vbCrLf & sFrame(I)
    'Run the current frame, it can be a bitmap to render,
    'or a sound file to play
    'Sound , think.wav
    '100,thinking0000.bmp
    If LCase(sFrameData(0)) = "sound" Then
        'play the sound file
        'MsgBox "Playing Sound"
        If CheckFile(App.Path & "\temp\" & sName & "\" & sFrameData(1)) = True Then
            PlayAWav App.Path & "\temp\" & sName & "\" & sFrameData(1)
        End If
    Else
        'render the bitmap
        'shape the display form to match the window region
        If CheckFile(App.Path & "\temp\" & sName & "\" & sFrameData(1)) = True Then
            'MsgBox "Doing Image"
            frmRender.Visible = True
            frmOffScreen.Picture1.Picture = LoadPicture(App.Path & "\temp\" & sName & "\" & sFrameData(1))
            frmOffScreen.Picture = frmOffScreen.Picture1.Picture
            frmOffScreen.Width = frmOffScreen.Picture1.Width
            frmOffScreen.Height = frmOffScreen.Picture1.Height
            SetRegion frmOffScreen, frmRender
            
            'wait the specified amount of time
            Wait CLng(sFrameData(0))
        End If
    End If
End If
Next i

DoEvents


Exit Sub
EH:
    ShowError Err.Number, Err.Description, "DoAnimation"
Exit Sub
End Sub
