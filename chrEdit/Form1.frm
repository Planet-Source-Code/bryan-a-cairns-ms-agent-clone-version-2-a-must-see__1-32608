VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "My Agent Charactor Editor"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog commdlg 
      Left            =   6720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":015C
            Key             =   ""
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02B8
            Key             =   ""
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0414
            Key             =   ""
            Object.Tag             =   "New Action"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AE8
            Key             =   ""
            Object.Tag             =   "New Frames"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C44
            Key             =   ""
            Object.Tag             =   "Remove Frames"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DA0
            Key             =   ""
            Object.Tag             =   "Add Sound"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11F4
            Key             =   ""
            Object.Tag             =   "Modify Frame Time"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1648
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1900
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3154
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5820
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Charactor"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Charactor"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Charactor"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Action"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Frames"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove Frames"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Action Sound"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Set Frame Display Time"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Move Frame Up"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Move Frame Down"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Play"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   15
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Action"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5295
      Left            =   2400
      TabIndex        =   3
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewCharactor 
         Caption         =   "New Charactor"
      End
      Begin VB.Menu mnuOpenCharactor 
         Caption         =   "Open Charactor"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveCharactor 
         Caption         =   "Save Charactor"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuSelNone 
         Caption         =   "Select None"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransColor 
         Caption         =   "Transparent Color"
      End
      Begin VB.Menu mnuActionPreview 
         Caption         =   "Action Preview"
      End
   End
   Begin VB.Menu mnuActionS 
      Caption         =   "Actions"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play Action"
      End
      Begin VB.Menu mnuStopAction 
         Caption         =   "Stop Action"
      End
      Begin VB.Menu mnuActionSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddSound 
         Caption         =   "Add Action Sound"
      End
      Begin VB.Menu mnuActionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewAction 
         Caption         =   "New Action"
      End
      Begin VB.Menu mnuActionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteAction 
         Caption         =   "Delete Action"
      End
   End
   Begin VB.Menu mnuFrameS 
      Caption         =   "Frames"
      Begin VB.Menu mnuAddFrames 
         Caption         =   "Add Frames"
      End
      Begin VB.Menu mnuDeleteFrames 
         Caption         =   "Delete Frames"
      End
      Begin VB.Menu mnuFrameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move Up"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move Down"
      End
      Begin VB.Menu mnuFrameSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModDisplayTimes 
         Caption         =   "Modify Display Times"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuShowHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OldAction As String

Private Sub Form_Load()
DoNew
End Sub

Private Sub Form_Resize()
On Error Resume Next
With ListView1
    .Height = (Me.Height - .Top) - 950
End With
With ListView2
    .Height = (Me.Height - .Top) - 950
    .Width = (Me.Width - .Left) - 120
    .ColumnHeaders(2).Width = (.Width - .ColumnHeaders(2).Left) - 360
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim Ret As Object
RmTree App.Path & "\temp\"
For Each Ret In Forms
    Unload Ret
Next Ret
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If OldAction <> "" Then
    WriteFramesFile OldAction
End If
LoadFrames Item.Text
OldAction = Item.Text
If ListView2.ListItems.Count > 0 Then DisplayCurrentFrame
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
DisplayCurrentFrame
End Sub

Private Sub mnuActionPreview_Click()
frmRender.Show
End Sub

Private Sub mnuAddFrames_Click()
AddFrames
End Sub

Private Sub mnuAddSound_Click()
AddActionSound
End Sub

Private Sub mnuDeleteAction_Click()
If ListView1.SelectedItem Is Nothing Then Exit Sub
Dim Ret
Ret = MsgBox("Are you sure you wish to delete the selected action?", vbYesNo, "Delete Action")
If Ret <> vbYes Then Exit Sub
DeleteAction ListView1.SelectedItem.Text

End Sub

Private Sub mnuDeleteFrames_Click()
DeleteFrame
End Sub

Private Sub mnuModDisplayTimes_Click()
Dim i As Integer
Dim s As String
Dim LTime As Integer
s = InputBox("Please enter the display time for the selected frames (in milliseconds)", "Frame Display Times", "100")
If IsNumeric(s) = False Then
    ShowInfo "Please enter numbers only!", "User Error"
    Exit Sub
End If

LTime = CInt(s)

If LTime < 50 Then LTime = 50
For i = 1 To ListView2.ListItems.Count
If ListView2.ListItems(i).Selected = True Then
    If LCase(ListView2.ListItems(i)) <> "sound" Then
        ListView2.ListItems(i).Text = LTime
    End If
End If
Next i
End Sub

Private Sub mnuMoveDown_Click()
MoveDown
End Sub

Private Sub mnuMoveUp_Click()
MoveUp
End Sub

Private Sub mnuNewAction_Click()
Dim s As String
s = InputBox("Please enter a name for this action.", "New Action", "New")
If s <> "" Then
    AddAction s
End If
End Sub


Private Sub MoveUp()
On Error GoTo EH
Dim ITMX As ListItem

If ListView2.SelectedItem.Index = 1 Then Exit Sub
Set ITMX = ListView2.ListItems.Add(ListView2.SelectedItem.Index - 1, , ListView2.SelectedItem.Text, , ListView2.SelectedItem.SmallIcon)
ITMX.SubItems(1) = ListView2.SelectedItem.SubItems(1)
If ListView2.SelectedItem.Index = ListView2.ListItems.Count Then
ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
Set ListView2.SelectedItem = ListView2.ListItems(ListView2.SelectedItem.Index - 1)
Else
ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
Set ListView2.SelectedItem = ListView2.ListItems(ListView2.SelectedItem.Index - 2)
End If
Exit Sub
EH:
Exit Sub
End Sub

Private Sub MoveDown()
On Error GoTo EH
Dim ITMX As ListItem

If ListView2.SelectedItem.Index = ListView2.ListItems.Count Then Exit Sub

Set ITMX = ListView2.ListItems.Add(ListView2.SelectedItem.Index + 2, , ListView2.SelectedItem.Text, , ListView2.SelectedItem.SmallIcon)
ITMX.SubItems(1) = ListView2.SelectedItem.SubItems(1)
ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
Set ListView2.SelectedItem = ListView2.ListItems(ListView2.SelectedItem.Index + 1)
Exit Sub
EH:
Exit Sub
End Sub

Private Sub mnuNewCharactor_Click()
DoNew
End Sub

Private Sub mnuOpenCharactor_Click()
DoOpen
End Sub

Private Sub mnuPlay_Click()
LoopAction
End Sub

Private Sub mnuSaveCharactor_Click()
DoSave
End Sub

Private Sub mnuSelAll_Click()
Dim i As Integer

For i = 1 To ListView2.ListItems.Count
    ListView2.ListItems(i).Selected = True
Next i
End Sub

Private Sub mnuSelNone_Click()
Dim i As Integer

For i = 1 To ListView2.ListItems.Count
    ListView2.ListItems(i).Selected = False
Next i
End Sub

Private Sub mnuShowHelp_Click()
ShowInfo "Sorry no help file yet", "Opps"
Exit Sub
ShowHelp 10
End Sub

Private Sub mnuStopAction_Click()
bCancelRender = True
End Sub

Private Sub mnuTransColor_Click()
On Error GoTo EH
Form1.commdlg.Filename = ""
Form1.commdlg.Color = lColor
Form1.commdlg.CancelError = True
Form1.commdlg.ShowColor

lColor = Form1.commdlg.Color
DisplayCurrentFrame
Exit Sub
EH:
If Err <> cdlCancel Then
ShowError Err.Number, Err.Description, "Selecting Color"
End If
Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1 'new
    mnuNewCharactor_Click
Case Is = 2 'open
    mnuOpenCharactor_Click
Case Is = 3 'save
    mnuSaveCharactor_Click
Case Is = 5 'add action
    mnuNewAction_Click
Case Is = 6 'add frame
    mnuAddFrames_Click
Case Is = 7 'remove frame
    mnuDeleteFrames_Click
Case Is = 8 'add sound
    mnuAddSound_Click
Case Is = 9 'mod time
    mnuModDisplayTimes_Click
Case Is = 11 'move up
    mnuMoveUp_Click
Case Is = 12 'move down
    mnuMoveDown_Click
Case Is = 14 'play
    mnuPlay_Click
Case Is = 15 'stop
    mnuStopAction_Click
Case Is = 17 'help
    mnuShowHelp_Click
End Select
End Sub
