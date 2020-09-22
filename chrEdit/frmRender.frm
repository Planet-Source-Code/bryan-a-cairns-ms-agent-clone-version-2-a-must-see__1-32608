VERSION 5.00
Begin VB.Form frmRender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Current Action"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2190
   Icon            =   "frmRender.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2190
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
TransparentBlt Picture1.hDC, 0, 0, Picture1.Width / 15, Picture1.Height / 15, pichidden.hDC, 0, 0, pichidden.Width / 15, pichidden.Height / 15, lColor
End Sub

Private Sub Form_Load()
SetWindowWord Me.hwnd, -8, Form1.hwnd
bRenderOpen = True
DisplayCurrentFrame
End Sub

Private Sub Form_Unload(Cancel As Integer)
bRenderOpen = False
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
Me.Width = Picture1.Width
Me.Height = Picture1.Height + 360
End Sub
