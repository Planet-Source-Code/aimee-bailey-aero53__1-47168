VERSION 5.00
Object = "{39A005FB-8F21-47F0-AD35-C52EE2FDC141}#13.0#0"; "XP_Command1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00008000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ScoreBoard"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3960
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XP_command_but.XPCommand XPCommand1 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "&Back"
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   49152
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pos"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Score"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function LoadScores()
'On Error GoTo err
Open App.Path & "\scores.dat" For Input As #1
Do Until EOF(1)
DoEvents
ListView1.ListItems.Add , , ListView1.ListItems.Count
Input #1, a$
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , a$
Input #1, a$
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , a$
Loop
err:
Close #1
End Function
Public Function AddScore(xname As String, score As Integer)
LoadScores
ListView1.ListItems.Add , , "1"
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xname
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , score
ReOrder
SaveScores
End Function

Public Function ReOrder()
For i = 1 To ListView1.ListItems.Count - 1
ListView1.ListItems(i).Text = i
Next i
End Function

Public Function SaveScores()
Open App.Path & "\scores.dat" For Output As #1
For i = 1 To ListView1.ListItems.Count - 1
Print #1, ListView1.ListItems(i + 1).ListSubItems(1).Text
Print #1, ListView1.ListItems(i + 1).ListSubItems(2).Text
Next i
Close #1
End Function
