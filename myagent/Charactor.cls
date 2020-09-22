VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Charactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub StartCharactor(sFile As String)
On Error GoTo EH
bCancel = False
bLoaded = True
If LoadCHRFile(sFile) = False Then
    DoClose
    ShowInfo "Could not start Charactor!", "Error"
End If
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Starting Charactor"
Exit Sub
End Sub

Public Sub StopAction()
bCancel = True
DoEvents
 Wait 500
DoAnimation "idle"
End Sub
Public Sub DoAction(sName As String)
On Error GoTo EH
Dim bOK As Boolean
If bLoaded = False Then
    ShowInfo "No Charactor Loaded!", "Error"
    Exit Sub
End If
bCancel = True
DoEvents
Wait 300
bCancel = False

DoAnimation sName
'revert back to "Idle" action when finished
'Idle is the ONLY default action you MUST have
'in your chractor file
 Wait 500
DoAnimation "idle"

Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Starting Charactor"
Exit Sub
End Sub

Public Sub CloseCharactor()
On Error GoTo EH
DoClose
Exit Sub
EH:
    ShowError Err.Number, Err.Description, "Closing Charactor"
Exit Sub
End Sub

Public Sub ShowText(sTitle As String, sText As String, bCloseButton As Boolean, CloseTime As Long, sIcon As String)
On Error GoTo EH
If bLoaded = False Then Exit Sub
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(frmRender.hwnd, WinRect)
BalloonXY.X = (WinRect.Left * Screen.TwipsPerPixelX)
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY

Set fBalloon = New frmBalloon
'icon styles are...
'"i" 'The "i" icon, XP-style (default)
'"i9" 'The "i" icon, 9x/Me-style
'"x" 'The "x" icon, XP-style
'"x9" 'The "x" icon, 9x/Me-style
'"!" 'The "!" icon, XP-style
'"!9" 'The "!" icon, 9x-style
'Use no icon
    
fBalloon.SetBalloon sTitle, sText, BalloonXY.X, BalloonXY.Y, sIcon, bCloseButton, CloseTime
    fBalloon.Move frmRender.Left - (frmRender.Width / 2), (frmRender.Top - fBalloon.Height)
fBalloon.Show , frmRender
frmRender.SetFocus
Exit Sub
EH:
    ShowError Err.Description, Err.Number, "ShowText"
Exit Sub
End Sub
