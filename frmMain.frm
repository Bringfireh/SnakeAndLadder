VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake and Ladder Game."
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   11820
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   99
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   98
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   97
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   96
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   95
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   94
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   93
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   92
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   91
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   90
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   89
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   88
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   87
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   86
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   85
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   84
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   83
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   82
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   81
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   80
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   79
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   78
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   77
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   76
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   75
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   74
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   73
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   72
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   71
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   70
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   69
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   68
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   67
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   66
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   65
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   64
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   63
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   62
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   61
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   60
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   59
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   58
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   57
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   56
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   55
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   54
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   53
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   52
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   51
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   50
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   49
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   48
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   47
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   46
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   45
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   44
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   43
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   42
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   41
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   40
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   39
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   38
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   37
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   36
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   35
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   34
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   33
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   32
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   31
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   30
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   29
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   28
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   27
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   26
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   25
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   24
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   23
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   22
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   21
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   20
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   19
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   18
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   17
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   16
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   15
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   14
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   13
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   12
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   11
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   10
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   8
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   7
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   6
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   8760
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H00008080&
         Caption         =   "Play"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LRESULT 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   240
         TabIndex        =   104
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblDie 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Player1-Your turn"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   120
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   8535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuplayer1 
         Caption         =   "Player1"
         Begin VB.Menu mnuRed 
            Caption         =   "Red"
         End
      End
      Begin VB.Menu mnuPLayer2 
         Caption         =   "Player2"
         Begin VB.Menu mnuBlue1 
            Caption         =   "Blue"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuIndex 
         Caption         =   "Index"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Snake N' Ladder Game"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim n As Integer
Dim p1 As Integer
Dim p2 As Integer
Dim d As Integer
Private Sub cmdPlay_Click()
    Randomize Timer
    For n = 1 To 6
    x = Int(Rnd * 6) + 1
    lblDie.Caption = x
    Next
        
    If lblPlayer.Caption = "Player1-Your turn" Then
        If p1 = 0 Then
            If Val(lblDie.Caption) = 6 Then
                p1 = 1
                Command2(0).Visible = True
                Command2(0).BackColor = vbRed
                Changelabel
                Exit Sub
            ElseIf Val(lblDie.Caption) <> 6 Then
                Changelabel
                Exit Sub
            End If
        ElseIf p1 <> 0 Then
            checkLabel
            checknewpos
            Exit Sub
        End If
            
    ElseIf lblPlayer.Caption = "Player2-Your turn" Then
        If p2 = 0 Then
            If Val(lblDie.Caption) = 6 Then
                p2 = 1
                Command2(0).Visible = True
                Command2(0).BackColor = vbBlue
                Changelabel
            ElseIf Val(lblDie.Caption) <> 6 Then
                Changelabel
                Exit Sub
            End If
        ElseIf p2 <> 0 Then
            checkLabel
            checknewpos
            Exit Sub
        End If
    End If
        
        
            
  
    
        
End Sub

Private Function checknewpos()
    If Command2(99).Visible = True Then
    If lblPlayer.Caption = "Player1-Your turn" Then
    LRESULT.Caption = "Player2 You Win"
        mnuNew_Click
    Else
    LRESULT.Caption = "Player1-You Win"
        mnuNew_Click
    End If
    End If
End Function




Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then endgame
If KeyAscii = 107 Then cmdPlay_Click
End Sub

Private Sub Form_Load()
    p1 = 0
    p2 = 0
    lblPlayer.Caption = "Player1-Your turn"
    lblDie.Caption = ""
End Sub
 Private Sub checkLabel()
        If lblPlayer.Caption = "Player1-Your turn" Then
        Command2(p1 - 1).Visible = False
        p1 = p1 + Val(lblDie.Caption)
        CheckSNL1
        If p1 >= 100 Then p1 = 100
        Command2(p1 - 1).Visible = True: Command2(p1 - 1).BackColor = vbRed: If p2 <> 0 Then Command2(p2 - 1).Visible = True: Command2(p2 - 1).BackColor = vbBlue
        Changelabel
        Exit Sub
        ElseIf lblPlayer.Caption = "Player2-Your turn" Then
        Command2(p2 - 1).Visible = False
        p2 = p2 + Val(lblDie.Caption)
        CheckSNL2
        If p2 >= 100 Then p2 = 100
        Command2(p2 - 1).Visible = True: Command2(p2 - 1).BackColor = vbBlue: If p1 <> 0 Then Command2(p1 - 1).Visible = True: Command2(p1 - 1).BackColor = vbRed
        Changelabel
        Exit Sub
        End If
 End Sub
 Private Sub Changelabel()
    If lblPlayer.Caption = "Player1-Your turn" Then
        lblPlayer.Caption = "Player2-Your turn"
    Else
        lblPlayer.Caption = "Player1-Your turn"
    End If
 End Sub
 
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuIndex_Click()
    frmIndex.Show
End Sub

Private Sub mnuNew_Click()
    If MsgBox("Do you want to start a new game?", vbYesNo, "New Game") = vbNo Then endgame: Exit Sub
        For d = 0 To 99
        Command2(d).Visible = False
        Next
        LRESULT.Caption = ""
    Form_Load
End Sub
Private Sub CheckSNL1()
snakes:
    If p1 = 28 Then p1 = 9: MsgBox "moved from 28 back to 9", vbInformation, "SNAKE"
    If p1 = 35 Then p1 = 7: MsgBox "moved from 35 back to 7", vbInformation, "SNAKE"
    If p1 = 43 Then p1 = 24: MsgBox "moved from 43 back to 24", vbInformation, "SNAKE"
    If p1 = 52 Then p1 = 32: MsgBox "moved from 52 back to 32", vbInformation, "SNAKE"
    If p1 = 83 Then p1 = 57: MsgBox "moved from 83 back to 57", vbInformation, "SNAKE"
    If p1 = 86 Then p1 = 54: MsgBox "moved from 86 back to 54", vbInformation, "SNAKE"
    If p1 = 91 Then p1 = 69: MsgBox "moved from 91 back to 69", vbInformation, "SNAKE"
    If p1 = 99 Then p1 = 1: MsgBox "moved from 99 back to 1", vbInformation, "SNAKE"
ladders:
    If p1 = 17 Then p1 = 45: MsgBox "Moved from 17 to 45", vbInformation, "LADDER"
    If p1 = 20 Then p1 = 40: MsgBox "Moved from 20 to 40", vbInformation, "LADDER"
    If p1 = 33 Then p1 = 53: MsgBox "Moved from 33 to 53", vbInformation, "LADDER"
    If p1 = 55 Then p1 = 65: MsgBox "Moved from 55 to 65", vbInformation, "LADDER"
    If p1 = 63 Then p1 = 79: MsgBox "Moved from 63 to 79", vbInformation, "LADDER"
    If p1 = 87 Then p1 = 93: MsgBox "Moved from 87 to 93", vbInformation, "LADDER"
    
End Sub
Private Sub CheckSNL2()
snakes:
    If p2 = 28 Then p2 = 9: MsgBox "moved from 28 back to 9", vbInformation, "SNAKE"
    If p2 = 35 Then p2 = 7: MsgBox "moved from 35 back to 7", vbInformation, "SNAKE"
    If p2 = 43 Then p2 = 24: MsgBox "moved from 43 back to 24", vbInformation, "SNAKE"
    If p2 = 52 Then p2 = 32: MsgBox "moved from 52 back to 32", vbInformation, "SNAKE"
    If p2 = 83 Then p2 = 57: MsgBox "moved from 83 back to 57", vbInformation, "SNAKE"
    If p2 = 86 Then p2 = 54: MsgBox "moved from 86 back to 54", vbInformation, "SNAKE"
    If p2 = 91 Then p2 = 69: MsgBox "moved from 91 back to 69", vbInformation, "SNAKE"
    If p2 = 99 Then p2 = 1: MsgBox "moved from 99 back to 1", vbInformation, "SNAKE"
ladders:
    If p2 = 17 Then p2 = 45: MsgBox "Moved from 17 to 45", vbInformation, "LADDER"
    If p2 = 20 Then p2 = 40: MsgBox "Moved from 20 to 40", vbInformation, "LADDER"
    If p2 = 33 Then p2 = 53: MsgBox "Moved from 33 to 53", vbInformation, "LADDER"
    If p2 = 55 Then p2 = 65: MsgBox "Moved from 55 to 65", vbInformation, "LADDER"
    If p2 = 63 Then p2 = 79: MsgBox "Moved from 63 to 79", vbInformation, "LADDER"
    If p2 = 87 Then p2 = 93: MsgBox "Moved from 87 to 93", vbInformation, "LADDER"
    
End Sub
Private Function endgame()
   If MsgBox("Do you want to end the game?", vbYesNo, Me.Caption) = vbYes Then mnuExit_Click Else MsgBox "you have to start a new game any way.", vbInformation, Me.Caption: mnuNew_Click
End Function
