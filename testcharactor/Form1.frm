VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MS Agent Clone"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Stop Action"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Charactor"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run Action"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Charactor"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "D:\components\agentclone\actors\Bryan\bryan.chr"
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label5 
      Caption         =   "Step 4 - Close the Charactor when finished"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Step 3 - choose an Action from the list below"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Step 2 - Load the charactor file (see Read Me for file format)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Step 1 - If you have not already, compile the ""MyAgent"" Project included (Active X EXE project)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "MS Agent Clone in VB - make your own Charactors!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCharactor As Object

Private Sub Command1_Click()
On Error Resume Next
Set oCharactor = Nothing
Set oCharactor = CreateObject("MyAgent.Charactor")
oCharactor.StartCharactor Text1.Text
oCharactor.DoAction "idle"
End Sub

Private Sub Command2_Click()
'Simply call the action by name and the "MyAgent" object will play it
'but only if it exists in the charactor
    oCharactor.DoAction List1.Text
End Sub

Private Sub Command3_Click()
If oCharactor Is Nothing Then Exit Sub
oCharactor.CloseCharactor
End Sub

Private Sub Command4_Click()
oCharactor.stopaction
End Sub

Private Sub Form_Load()
'Text1.Text = App.Path & "\monitor.chr"
With List1
    .AddItem "idle"
    .AddItem "thinking"
    .AddItem "bow"
    .AddItem "LookLeft"
    .AddItem "LookRight"
    .AddItem "question"
    .AddItem "explain"
End With
List1.ListIndex = 0
Text1.Text = App.Path & "\bryan.chr"
End Sub
