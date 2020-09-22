Attribute VB_Name = "mod_Render"
Global bRenderOpen As Boolean
Global bCancelRender As Boolean
Public Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Integer, ByVal nYOriginDest As Integer, ByVal nWidthDest As Integer, ByVal nHeightDest As Integer, ByVal hdcSrc As Long, ByVal nXOriginSrc As Integer, ByVal nYOriginSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Long) As Boolean
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long


Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
'wait for a certain amount of time  - in milliseconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait '* 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
    DoEvents
        If bCancelRender = True Then
            Exit Do
        End If
        DoEvents
    Loop
End Function

Public Sub LoopAction()
bCancelRender = False
Dim i As Integer
If Form1.ListView2.ListItems.Count = 0 Then Exit Sub
If bRenderOpen = False Then
    frmRender.Show
    DoEvents
End If
Do
If bCancelRender = True Then Exit Do
For i = 1 To Form1.ListView2.ListItems.Count
    Form1.ListView2.ListItems(i).Selected = False
Next i

For i = 1 To Form1.ListView2.ListItems.Count
StopAllWavs
Form1.ListView2.ListItems(i).Selected = True
    If bCancelRender = True Then Exit For
    If LCase(Form1.ListView2.ListItems(i).Text) = "sound" Then
    'play the sound file
        PlayAWav App.Path & "\temp\" & Form1.ListView1.SelectedItem.Text & "\" & Form1.ListView2.ListItems(i).SubItems(1)
    Else
        DisplayCurrentFrame
        'wait for the specified amount of time
        Wait CLng(Form1.ListView2.ListItems(i).Text)
    End If
Next i
DoEvents
Loop
End Sub

Public Sub DisplayCurrentFrame()
On Error GoTo EH
If bRenderOpen = False Then Exit Sub
If Form1.ListView2.SelectedItem Is Nothing Then Exit Sub

Dim sFile As String

sFile = App.Path & "\temp\" & Form1.ListView1.SelectedItem.Text & "\" & Form1.ListView2.SelectedItem.SubItems(1)
frmRender.Picture1.Picture = LoadPicture(sFile)
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Displaying Frame"
Exit Sub
End Sub

