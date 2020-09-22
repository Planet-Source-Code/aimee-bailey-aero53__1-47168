VERSION 5.00
Object = "{39A005FB-8F21-47F0-AD35-C52EE2FDC141}#13.0#0"; "XP_Command1.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aero 53 (By Steve Bailey)"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Automation 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin XP_command_but.XPCommand XPCommand5 
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "E&xit"
   End
   Begin XP_command_but.XPCommand XPCommand3 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "&New Game"
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin XP_command_but.XPCommand XPCommand7 
         Height          =   315
         Left            =   5160
         TabIndex        =   73
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Finish"
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   2655
         TabIndex        =   72
         Top             =   3840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin XP_command_but.XPCommand XPCommand6 
         Height          =   315
         Left            =   5160
         TabIndex        =   71
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "&Next Round"
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   5520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   69
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   70
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   4920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   68
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   65
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   66
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   64
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   3120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   62
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   60
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   57
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   58
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   56
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   54
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   5520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   51
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   52
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   4920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   49
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   50
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   4320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   47
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   48
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   45
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   46
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   3120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   43
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   41
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   42
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1920
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   39
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   40
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1320
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   38
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   720
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   35
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   36
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   6120
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   5880
         Picture         =   "Form1.frx":024A
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   5640
         Picture         =   "Form1.frx":0494
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   5400
         Picture         =   "Form1.frx":06DE
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   5160
         Picture         =   "Form1.frx":0928
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4920
         Picture         =   "Form1.frx":0B72
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   4680
         Picture         =   "Form1.frx":0DBC
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4440
         Picture         =   "Form1.frx":1006
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4200
         Picture         =   "Form1.frx":1250
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   6120
         Picture         =   "Form1.frx":149A
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   5880
         Picture         =   "Form1.frx":16E4
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   22
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   5640
         Picture         =   "Form1.frx":192E
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   21
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   5400
         Picture         =   "Form1.frx":1B78
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   20
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   5160
         Picture         =   "Form1.frx":1DC2
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   19
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4920
         Picture         =   "Form1.frx":200C
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   18
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   4680
         Picture         =   "Form1.frx":2256
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4440
         Picture         =   "Form1.frx":24A0
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4200
         Picture         =   "Form1.frx":26EA
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3960
         Picture         =   "Form1.frx":2934
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox star2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3960
         Picture         =   "Form1.frx":2B7E
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   5160
         Top             =   2160
      End
      Begin XP_command_but.XPCommand XPCommand2 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Call"
      End
      Begin XP_command_but.XPCommand XPCommand1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "Hit Me"
      End
      Begin VB.PictureBox player 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   3120
         Width           =   495
         Begin VB.Label player1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox computer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   1
         Top             =   480
         Width           =   495
         Begin VB.Label computer1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "|||||"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   75
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Possible Opposition Count: 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   74
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "My Count: 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   34
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Round:  0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To Aero53!!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   6255
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   6360
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Player"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3255
      End
   End
   Begin XP_command_but.XPCommand XPCommand4 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GameStopped As Integer
Public CardPos As Integer
Public ComWinCount As Integer
Public PlayerWinCount As Integer
Public playerpos As Integer
Public round As Integer
Public score As Integer

Private Sub Automation_Timer()
On Error Resume Next
Dim x As Integer
x = False
Timer1.Enabled = GameStopped
If GameStopped = False Then
    x = True
    For i = 0 To 9
        If i <= ComWinCount Then star1(i).Visible = True Else star1(i).Visible = False
        If i <= PlayerWinCount Then star2(i).Visible = True Else star2(i).Visible = False
        
    Next i
    Label4.Caption = "Round: " & round
    Label6.Visible = True
    Label7.Caption = "Score: " & score
    Else
    Label6.Visible = False
End If

XPCommand1.Visible = x
XPCommand2.Visible = x
End Sub

Private Sub Form_Load()
GameStopped = True
End Sub

Private Sub Timer1_Timer()
If Label3.Caption = "Welcome To Aero53!!!" Then
Label3.Caption = "The game of chance!"
ElseIf Label3.Caption = "The game of chance!" Then
Label3.Caption = "And Twisting Fun!"
ElseIf Label3.Caption = "And Twisting Fun!" Then
Label3.Caption = "Press New Game To Start!!!"
ElseIf Label3.Caption = "Press New Game To Start!!!" Then
Label3.Caption = "Welcome To Aero53!!!"
End If
For i = 0 To 9
        If i <= ComWinCount Then star1(i).Visible = True Else star1(i).Visible = False
        If i <= PlayerWinCount Then star2(i).Visible = True Else star2(i).Visible = False
        
    Next i
End Sub

Private Sub XPCommand1_Click()
Dim x As Integer
playerpos = playerpos + 1
player(playerpos).Visible = True
computer(playerpos).Visible = True
Randomize
computer1(playerpos).Tag = Int(Rnd * 20)
Randomize
player1(playerpos).Caption = Int(Rnd * 20)
For i = 0 To playerpos
x = x + player1(i).Caption
Next i
Label5.Caption = "My Count: " & x
Randomize
Label6.Caption = "Possible Opposition Count: " & Int(Rnd * Int((playerpos + 1) * 20))
End Sub

Private Sub XPCommand2_Click()
Dim p_score As Integer
Dim c_score As Integer
Dim p_diff As Integer
Dim c_diff As Integer
Picture2.Visible = True


For i = 0 To playerpos
c_score = c_score + computer1(i).Tag
Next i

For a = 0 To playerpos
p_score = p_score + player1(a).Caption
Next a

If p_score < 53 Then
    p_diff = 53 - p_score
ElseIf p_score > 53 Then
    p_diff = p_score - 53
End If

If c_score < 53 Then
    c_diff = 53 - c_score
ElseIf c_score > 53 Then
    c_diff = c_score - 53
End If


If c_diff < p_diff Then
    y = "Computer Wins!!"
    ComWinCount = ComWinCount + 1
Else
    y = Label2.Caption & " Wins!!"
    PlayerWinCount = PlayerWinCount + 1
    score = score + (Int(p_score / playerpos) * 100)

End If
x = "Computer: " & c_score & "  " & Label2.Caption & ": " & p_score
x = x & vbCrLf & y

Label3.Caption = x
XPCommand6.Visible = True

If Label4.Caption = "Round: 9" Then
XPCommand7.Visible = True
If PlayerWinCount > ComWinCount Then
    Label3.Caption = "!!! End Of Match !!!" & vbCrLf & Label2.Caption & " Wins!!! " & PlayerWinCount & " / " & ComWinCount
    Else
    Label3.Caption = "!!! End Of Match !!!" & "The Computer Wins!!! " & ComWinCount & " / " & PlayerWinCount
End If
Dim xx As Integer
If score > 7000 Then
    xx = MsgBox("Enter score onto scoreboard?", vbYesNo, "Aero53")
    If xx = vbYes Then
    Form3.Show
    Form3.AddScore Label2.Caption, score
    End If
End If
End If

End Sub

Private Sub XPCommand3_Click()
Dim xname As String
Dim st As Integer
xname = InputBox("Please enter your name:", "New Game!", "Player1")
st = MsgBox("Ready to start?", vbYesNo)
If st = vbNo Then Exit Sub
Label2.Caption = xname
GameStopped = False
Randomize
computer1(0).Tag = Int(Rnd * 20)
computer1(0).Caption = "|||||"
Randomize
player1(0).Caption = Int(Rnd * 20)

Label5.Caption = "My Count: " & player1(0).Caption
Label3.Caption = "LETS ROLL!!!"
round = 1
Randomize
Label6.Caption = "Possible Opposition Count: " & Int(Rnd * Int((playerpos + 1) * 20))
score = 0
End Sub

Private Sub XPCommand4_Click()
Form2.Show vbModal
End Sub

Private Sub XPCommand5_Click()
Dim x As Integer
x = MsgBox("Are you sure you want to leave :(", vbYesNo, "Aero53")
If x = vbYes Then End
End Sub

Private Sub XPCommand6_Click()
XPCommand6.Visible = False
Picture2.Visible = False
Label3.Caption = "LETS ROLL!!!"
round = round + 1

For i = 0 To 9
computer1(i).Tag = 0
player1(i).Caption = 0
If i >= 1 Then
    computer(i).Visible = False
    player(i).Visible = False
End If
Next i
Label5.Caption = "My Count: " & player1(0).Caption
CardPos = 0
playerpos = 0
End Sub

Private Sub XPCommand7_Click()
GameStopped = True
XPCommand7.Visible = False
XPCommand6.Visible = False
Picture2.Visible = False
Label3.Caption = "Welcome To Aero53!!!"
round = 0

For i = 0 To 9
computer1(i).Tag = 0
player1(i).Caption = 0
If i >= 1 Then
    computer(i).Visible = False
    player(i).Visible = False
End If
Next i
Label5.Caption = "My Count: 0"
CardPos = 0
playerpos = 0
Label2.Caption = "Player1"
score = 0
End Sub
