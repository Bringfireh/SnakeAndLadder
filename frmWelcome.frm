VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome to Snake and Ladder game"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Height          =   1455
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   4935
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Welcome to the snake and ladder game"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label Label2 
            Caption         =   "Use K to play the game or Click play "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlayGame 
         Caption         =   "Play game"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "SNAKE AND LADDER GAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdPlayGame_Click()
    Unload Me
    Load frmMain
    frmMain.Show
End Sub
