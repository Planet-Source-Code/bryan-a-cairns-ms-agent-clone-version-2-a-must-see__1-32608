VERSION 5.00
Begin VB.Form frmRender 
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'set the height and width to 150 x 150
'then set the window to the topmost
Me.Width = 2250
Me.Height = 2250
OnTop Me, True
frmOffScreen.Show
frmOffScreen.Move 0 - (frmOffScreen.Width + 2000), 0
Me.Visible = False
Me.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
If hRgn Then DeleteObject hRgn
If bShowBallon = True Then
Unload fBalloon
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bShowBallon = True Then
    Unload fBalloon
    End If
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show a menu or some other nonsence here
End Sub
