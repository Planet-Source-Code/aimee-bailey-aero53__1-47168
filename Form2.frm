VERSION 5.00
Object = "{39A005FB-8F21-47F0-AD35-C52EE2FDC141}#13.0#0"; "XP_Command1.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Readme.txt"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XP_command_but.XPCommand XPCommand1 
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "< &Back"
   End
   Begin VB.TextBox Text1 
      Height          =   5655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open App.Path & "\readme.txt" For Input As #1
Do Until EOF(1)
DoEvents
Input #1, a$
Text1.Text = Text1.Text & a$ & vbCrLf
Loop
Close #1
End Sub

Private Sub XPCommand1_Click()
Unload Me
End Sub
