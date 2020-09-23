VERSION 5.00
Begin VB.Form Monopoly 
   BackColor       =   &H00C0FFCF&
   BorderStyle     =   0  'None
   Caption         =   "Online Monopoly"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Monopoly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   600
   End
   Begin VB.PictureBox picWaiting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   -4200
      Picture         =   "Monopoly.frx":08CA
      ScaleHeight     =   5625
      ScaleWidth      =   7905
      TabIndex        =   117
      Top             =   10320
      Width           =   7935
      Begin VB.OptionButton optRules 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Short Game Rules"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   146
         Top             =   4200
         Width           =   2895
      End
      Begin VB.OptionButton optRules 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Standard Rules"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   145
         Top             =   3840
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "start"
         Height          =   375
         Left            =   6480
         Picture         =   "Monopoly.frx":19992
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   4800
         Width           =   1215
      End
      Begin VB.PictureBox picWPInfoBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   3840
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   136
         Top             =   3240
         Width           =   3855
         Begin VB.Label lblWReady 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   143
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgWPToken 
            Height          =   615
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblWPJoinName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   137
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox picWPInfoBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   3840
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   134
         Top             =   2520
         Width           =   3855
         Begin VB.Label lblWReady 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   142
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgWPToken 
            Height          =   615
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblWPJoinName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   135
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox picWPInfoBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   3840
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   132
         Top             =   1800
         Width           =   3855
         Begin VB.Label lblWReady 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   140
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgWPToken 
            Height          =   615
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblWPJoinName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   133
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox picWPInfoBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   3840
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   130
         Top             =   3960
         Width           =   3855
         Begin VB.Label lblWReady 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   141
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblWPJoinName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   131
            Top             =   0
            Width           =   3015
         End
         Begin VB.Image imgWPToken 
            Height          =   615
            Index           =   4
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdWExit 
         Caption         =   "exit"
         Height          =   375
         Left            =   240
         Picture         =   "Monopoly.frx":32A5A
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdReady 
         Caption         =   "Ready"
         Height          =   375
         Left            =   6480
         Picture         =   "Monopoly.frx":4BB22
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Game Information"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   240
         TabIndex        =   119
         Top             =   240
         Width           =   3495
         Begin VB.Label lblWGameTitle 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   840
            TabIndex        =   123
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblWGameNo 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   840
            TabIndex        =   122
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Title :  "
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. :  "
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Player Information"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   3840
         TabIndex        =   118
         Top             =   240
         Width           =   3855
         Begin VB.Label lblWPName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   960
            TabIndex        =   127
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label lblWPID 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   960
            TabIndex        =   126
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name :  "
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ID : "
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label lblLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   139
         Top             =   5280
         Width           =   7935
      End
      Begin VB.Image imgWToken 
         Height          =   615
         Index           =   0
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   8415
      Left            =   -8160
      ScaleHeight     =   8385
      ScaleWidth      =   11145
      TabIndex        =   85
      Top             =   9600
      Visible         =   0   'False
      Width           =   11175
      Begin VB.PictureBox picStatusDeed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   4
         Left            =   8400
         ScaleHeight     =   6945
         ScaleWidth      =   2625
         TabIndex        =   95
         Top             =   600
         Width           =   2655
         Begin VB.Label lblNW 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   162
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label lblAssets 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   161
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Net worth : "
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   160
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Assets : "
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   159
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label lblStatusPlayerName 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   2655
         End
         Begin VB.Image imgPlayerDeed4 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.PictureBox picStatusDeed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   3
         Left            =   5640
         ScaleHeight     =   6945
         ScaleWidth      =   2625
         TabIndex        =   93
         Top             =   600
         Width           =   2655
         Begin VB.Label lblNW 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   158
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label lblAssets 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   157
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Net worth : "
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   156
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Assets : "
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   155
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Image imgPlayerDeed3 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblStatusPlayerName 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   94
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox picStatusDeed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   2
         Left            =   2880
         ScaleHeight     =   6945
         ScaleWidth      =   2625
         TabIndex        =   90
         Top             =   600
         Width           =   2655
         Begin VB.Label lblNW 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   154
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label lblAssets 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   153
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Net worth : "
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   152
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Assets : "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   151
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label lblStatusPlayerName 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   2655
         End
         Begin VB.Image imgPlayerDeed2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdCloseStatus 
         BackColor       =   &H00FF8080&
         Caption         =   "close"
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   7680
         Width           =   1815
      End
      Begin VB.PictureBox picStatusDeed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   6975
         Index           =   1
         Left            =   120
         ScaleHeight     =   6945
         ScaleWidth      =   2625
         TabIndex        =   88
         Top             =   600
         Width           =   2655
         Begin VB.Label lblNW 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   150
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label lblAssets 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   149
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Net worth : "
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   148
            Top             =   6600
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Assets : "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   147
            Top             =   6240
            Width           =   1215
         End
         Begin VB.Label lblStatusPlayerName 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   91
            Top             =   0
            Width           =   2655
         End
         Begin VB.Image imgPlayerDeed1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdPlayerStatus 
         BackColor       =   &H00FF8080&
         Caption         =   "players"
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "status"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   87
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.PictureBox picToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   1680
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   72
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Image imgToken 
         Height          =   495
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   3120
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   75
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Image imgToken 
         Height          =   495
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   3840
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   74
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Image imgToken 
         Height          =   495
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   2400
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   73
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Image imgToken 
         Height          =   495
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picMakeProposal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   8535
      Left            =   -7560
      ScaleHeight     =   8505
      ScaleWidth      =   9825
      TabIndex        =   63
      Top             =   9000
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CommandButton cmdDoneTrade 
         BackColor       =   &H00FF8080&
         Caption         =   "done"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdCounterTrade 
         BackColor       =   &H00FF8080&
         Caption         =   "Counter"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdRejectTrade 
         BackColor       =   &H00FF8080&
         Caption         =   "reject"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdAcceptTrade 
         BackColor       =   &H00FF8080&
         Caption         =   "accept"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   7920
         Width           =   1335
      End
      Begin VB.TextBox txtTradeAmount2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8400
         MaxLength       =   5
         TabIndex        =   99
         Text            =   "0"
         Top             =   6840
         Width           =   1335
      End
      Begin VB.TextBox txtTradeAmount1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         MaxLength       =   5
         TabIndex        =   97
         Text            =   "0"
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton cmdProposeTrade 
         Caption         =   "propose"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelTrade 
         Caption         =   "cancel"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   7920
         Width           =   1335
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   5640
         ScaleHeight     =   2145
         ScaleWidth      =   4065
         TabIndex        =   68
         Top             =   4560
         Width           =   4095
         Begin VB.Image imgTradingOutJail2 
            Height          =   495
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Image imgTradingCard2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2145
         ScaleWidth      =   4065
         TabIndex        =   67
         Top             =   4560
         Width           =   4095
         Begin VB.Image imgTradingOutJail1 
            Height          =   495
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Image imgTradingCard1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   6360
         ScaleHeight     =   3585
         ScaleWidth      =   3345
         TabIndex        =   66
         Top             =   840
         Width           =   3375
         Begin VB.Image imgTradeOutJail2 
            Height          =   495
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   3000
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Image imgTradeCard2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   3345
         TabIndex        =   65
         Top             =   840
         Width           =   3375
         Begin VB.Image imgTradeOutJail1 
            Height          =   495
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   3000
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Image imgTradeCard1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Label lblTradeAnounce 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   107
         Top             =   7320
         Width           =   9615
      End
      Begin VB.Label lblTraderName2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6360
         TabIndex        =   106
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblTraderName1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash : "
         Height          =   255
         Left            =   7560
         TabIndex        =   100
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash : "
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   6840
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   1380
         Left            =   4320
         Picture         =   "Monopoly.frx":64BEA
         Top             =   4920
         Width           =   1200
      End
      Begin VB.Image imgTradeCardPreview 
         Height          =   3255
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "proposal"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   0
         Width           =   9975
      End
   End
   Begin VB.PictureBox picSelectTrader 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   -2400
      ScaleHeight     =   3465
      ScaleWidth      =   4185
      TabIndex        =   56
      Top             =   8400
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Height          =   375
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdPlayerName 
         BackColor       =   &H0080FF80&
         Height          =   375
         Index           =   3
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdPlayerName 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Index           =   4
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdPlayerName 
         BackColor       =   &H00FF8080&
         Height          =   375
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancelTrading 
         BackColor       =   &H00FF8080&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblTrade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "who do you want to trade with?"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.PictureBox picMortgage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   -9840
      ScaleHeight     =   8265
      ScaleWidth      =   11145
      TabIndex        =   27
      Top             =   7800
      Visible         =   0   'False
      Width           =   11175
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FF8080&
         Caption         =   "close"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton cmdMortgageDC 
         BackColor       =   &H00FF8080&
         Caption         =   "mortgage"
         Height          =   375
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnmortgageDC 
         BackColor       =   &H00FF8080&
         Caption         =   "unmortgage"
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7680
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "mortgage/unmortgage"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   11175
      End
      Begin VB.Image imgCardPreview 
         Appearance      =   0  'Flat
         Height          =   3375
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image imgMortgageDeedCard 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.PictureBox picAuction 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   -6120
      ScaleHeight     =   6585
      ScaleWidth      =   6825
      TabIndex        =   31
      Top             =   7200
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Timer tmrAuctionDelay 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   4560
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$200"
         Height          =   615
         Index           =   6
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$100"
         Height          =   615
         Index           =   5
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$50"
         Height          =   615
         Index           =   4
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$20"
         Height          =   615
         Index           =   3
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$10"
         Height          =   615
         Index           =   2
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$5"
         Height          =   615
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton cmdBidCash 
         BackColor       =   &H00FF8080&
         Caption         =   "$1"
         Height          =   615
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5760
         Width           =   735
      End
      Begin VB.PictureBox picAuctionPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   3
         Left            =   3720
         ScaleHeight     =   825
         ScaleWidth      =   2745
         TabIndex        =   43
         Top             =   2520
         Width           =   2775
         Begin VB.Label lblCashBalance 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   45
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label lblAucName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.PictureBox picAuctionPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   2
         Left            =   3720
         ScaleHeight     =   825
         ScaleWidth      =   2745
         TabIndex        =   40
         Top             =   1560
         Width           =   2775
         Begin VB.Label lblCashBalance 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   42
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label lblAucName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.PictureBox picAuctionPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   3720
         ScaleHeight     =   825
         ScaleWidth      =   2745
         TabIndex        =   37
         Top             =   3480
         Width           =   2775
         Begin VB.Label lblAucName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label lblCashBalance 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   38
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.PictureBox picAuctionPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   3720
         ScaleHeight     =   825
         ScaleWidth      =   2745
         TabIndex        =   34
         Top             =   600
         Width           =   2775
         Begin VB.Label lblCashBalance 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   36
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label lblAucName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.Label lblCurrentBid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3600
         TabIndex        =   55
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bid :"
         Height          =   255
         Left            =   1920
         TabIndex        =   54
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblComment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   5160
         Width           =   6255
      End
      Begin VB.Image imgAuctionCard 
         Height          =   3735
         Left            =   240
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "auction"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   6855
      End
   End
   Begin VB.Timer tmrMoveToken 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer tmrDice 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picSideBar 
      Align           =   4  'Align Right
      BackColor       =   &H00C0FFCF&
      BorderStyle     =   0  'None
      Height          =   11520
      Left            =   7335
      ScaleHeight     =   11520
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdSetting 
         Caption         =   "setting"
         Height          =   375
         Left            =   360
         Picture         =   "Monopoly.frx":64DE2
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   11040
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Game Information"
         Height          =   1095
         Left            =   120
         TabIndex        =   114
         Top             =   120
         Width           =   3855
         Begin VB.Label lblGameTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Title :"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label lblGameID 
            BackStyle       =   0  'Transparent
            Caption         =   "ID :"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame frmPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "Current Player"
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         TabIndex        =   108
         Top             =   5400
         Width           =   3855
         Begin VB.Image imgDice2 
            Height          =   495
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   495
         End
         Begin VB.Image imgDice1 
            Height          =   495
            Left            =   1320
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblCurCash 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   840
            TabIndex        =   110
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label lblCurPName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   840
            TabIndex        =   109
            Top             =   360
            Width           =   2895
         End
         Begin VB.Image imgCurJail 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   120
            Picture         =   "Monopoly.frx":7DEAA
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Image imgCurPToken 
            Height          =   615
            Left            =   120
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image imgJailB 
            Height          =   1095
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Image imgJailA 
            Height          =   1095
            Left            =   120
            Stretch         =   -1  'True
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.Frame frmPInfo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Player Information"
         Height          =   1095
         Left            =   120
         TabIndex        =   78
         Top             =   1320
         Width           =   3855
         Begin VB.Label lblPName 
            BackStyle       =   0  'Transparent
            Caption         =   "Name: "
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label lblPID 
            BackStyle       =   0  'Transparent
            Caption         =   "ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmdMenu 
         Appearance      =   0  'Flat
         Caption         =   "exit"
         Height          =   375
         Left            =   2280
         MouseIcon       =   "Monopoly.frx":7E1CE
         MousePointer    =   99  'Custom
         Picture         =   "Monopoly.frx":7E4D8
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   11040
         Width           =   1215
      End
      Begin VB.PictureBox picPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   12
         Top             =   3960
         Width           =   3855
         Begin VB.Image imgJail 
            Appearance      =   0  'Flat
            Height          =   615
            Index           =   3
            Left            =   0
            Picture         =   "Monopoly.frx":15F61A
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblCashFlow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   24
            Top             =   0
            Width           =   975
         End
         Begin VB.Image imgPlayerToken 
            Height          =   615
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblPlayerCash 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblPlayerName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   19
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox picPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   11
         Top             =   3240
         Width           =   3855
         Begin VB.Image imgJail 
            Appearance      =   0  'Flat
            Height          =   615
            Index           =   2
            Left            =   0
            Picture         =   "Monopoly.frx":15F93E
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblCashFlow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   23
            Top             =   0
            Width           =   975
         End
         Begin VB.Image imgPlayerToken 
            Height          =   615
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblPlayerCash 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblPlayerName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   17
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox picPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   10
         Top             =   2520
         Width           =   3855
         Begin VB.Image imgJail 
            Appearance      =   0  'Flat
            Height          =   615
            Index           =   1
            Left            =   0
            Picture         =   "Monopoly.frx":15FC62
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblCashFlow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   22
            Top             =   0
            Width           =   975
         End
         Begin VB.Image imgPlayerToken 
            Height          =   615
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblPlayerCash 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblPlayerName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   15
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox picPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   9
         Top             =   4680
         Width           =   3855
         Begin VB.Image imgJail 
            Appearance      =   0  'Flat
            Height          =   615
            Index           =   4
            Left            =   0
            Picture         =   "Monopoly.frx":15FF86
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblCashFlow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   25
            Top             =   0
            Width           =   975
         End
         Begin VB.Image imgPlayerToken 
            Height          =   615
            Index           =   4
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblPlayerCash 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblPlayerName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   13
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Frame frmButton 
         BackColor       =   &H00C0FFC0&
         Height          =   2655
         Left            =   0
         TabIndex        =   1
         Top             =   8280
         Width           =   4095
         Begin VB.CommandButton cmdPayTax 
            Caption         =   "pay"
            Height          =   375
            Left            =   120
            Picture         =   "Monopoly.frx":1602AA
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdPayPercentTax 
            Caption         =   "tax"
            Height          =   375
            Left            =   1440
            Picture         =   "Monopoly.frx":179372
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdStatus 
            Caption         =   "status"
            Height          =   375
            Left            =   2040
            Picture         =   "Monopoly.frx":25A4B4
            Style           =   1  'Graphical
            TabIndex        =   111
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdUseCard 
            Caption         =   "use card"
            Height          =   375
            Left            =   1440
            Picture         =   "Monopoly.frx":27357C
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdPay 
            Caption         =   "pay"
            Height          =   375
            Left            =   1440
            Picture         =   "Monopoly.frx":28C644
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdBankrupt 
            Caption         =   "bankrupt"
            Height          =   375
            Left            =   2760
            Picture         =   "Monopoly.frx":36D786
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdBuild 
            Caption         =   "build"
            Height          =   375
            Left            =   360
            Picture         =   "Monopoly.frx":44E8C8
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdBuy 
            Caption         =   "buy"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            Picture         =   "Monopoly.frx":52FA0A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdTrade 
            Caption         =   "trade"
            Height          =   375
            Left            =   720
            Picture         =   "Monopoly.frx":610B4C
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDone 
            Caption         =   "done"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            Picture         =   "Monopoly.frx":629C14
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdSell 
            Caption         =   "sell"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            Picture         =   "Monopoly.frx":642CDC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdAuction 
            Caption         =   "auction"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Picture         =   "Monopoly.frx":723E1E
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdMortgage 
            Caption         =   "mortgage/unmortgage"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Picture         =   "Monopoly.frx":804F60
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1440
            Width           =   3855
         End
         Begin VB.CommandButton cmdRollDice 
            Caption         =   "roll dice"
            Height          =   375
            Left            =   120
            Picture         =   "Monopoly.frx":8E60A2
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   840
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picChatBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   1800
      ScaleHeight     =   2745
      ScaleWidth      =   7305
      TabIndex        =   81
      Top             =   6360
      Width           =   7335
      Begin VB.TextBox txtTblChatMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Top             =   360
         Width           =   7335
      End
      Begin VB.TextBox txtSendTableMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         MaxLength       =   80
         TabIndex        =   82
         Top             =   2160
         Width           =   7335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Chat Box"
         Height          =   375
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Image imgHotel 
      Height          =   255
      Index           =   0
      Left            =   960
      Picture         =   "Monopoly.frx":9C71E4
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgHouse2 
      Height          =   255
      Index           =   0
      Left            =   840
      Picture         =   "Monopoly.frx":9C73C8
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgHouse4 
      Height          =   255
      Index           =   0
      Left            =   1320
      Picture         =   "Monopoly.frx":9C75AC
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgHouse3 
      Height          =   255
      Index           =   0
      Left            =   1080
      Picture         =   "Monopoly.frx":9C7790
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgHouse1 
      Height          =   255
      Index           =   0
      Left            =   600
      Picture         =   "Monopoly.frx":9C7974
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDeedCard 
      Height          =   3255
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image imgCard 
      Height          =   1935
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Shape shpSlot 
      BackColor       =   &H80000000&
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgSlot 
      Appearance      =   0  'Flat
      Height          =   840
      Index           =   0
      Left            =   120
      MouseIcon       =   "Monopoly.frx":9C7B58
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "Monopoly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const slotSize = 1440
Const slotSizeW = slotSize * 4 / 6.5
Const startTop = 100
Const startLeft = 100

Public p As Integer
Public rollDiceDblCount As Integer
Public tempDice As Integer
Public numDice1 As Integer
Public numDice2 As Integer
Dim z As Integer

'##### Auction
Dim selectedCard As Integer
Dim timeDelay As Integer
Dim auctionTimeCount As Integer
Dim tempBider As Integer
Dim tempBidCard As Integer
'##### Trade
Dim tradingWith As Integer
Dim propertiesTrade1(1 To 20) As Integer
Dim propertiesTrade2(1 To 20) As Integer
Dim amountTrade As Currency
Dim outOfJailTradeA As Boolean
Dim outOfJailTradeB As Boolean

Dim tempNumber As Integer
Dim tempNumOfMove As Integer
Dim tempMoveDest As Integer
Dim tempMoveCheck As Boolean

Dim viewSlot As Integer
Dim loopX As Integer

Private Sub cmdSetting_Click()
    Load frmSetting
    frmSetting.Show
End Sub

Private Sub Form_Load()
    Set preActiveWindow = activeWindow
    Set activeWindow = Me
    frmMain.tmrMusic.Enabled = False
    tmrMusic.Enabled = True
    'On Error Resume Next
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Form_Resize

    picSideBar.Visible = False
    picAuction.Visible = False
    picChatBox.Visible = False
    picMortgage.Visible = False
    picSelectTrader.Visible = False
    picMakeProposal.Visible = False
    cmdReady.Visible = False
    cmdStart.Visible = False
    
    player(PlayerNumber).PID = PlayerID
    player(PlayerNumber).PName = PlayerName
    lblPID.Caption = "ID : " & PlayerID
    lblPName.Caption = "Name : " & PlayerName

    gameStarted = False
    Call rulesSetting
    Call viewWaitingWin
    Call loadDeedCardInfo
    frmMain.wskClient.SendData "status;enterGame;" & tableID & ";"
    DoEvents

    cmdStatus.Visible = False
    cmdTrade.Visible = False
    cmdBuy.Visible = False
    cmdSell.Visible = False
    cmdDone.Visible = False
    cmdPay.Visible = False
    cmdPayTax.Visible = False
    cmdPayPercentTax.Visible = False
    cmdUseCard.Visible = False
    cmdRollDice.Visible = False
    cmdAuction.Visible = False
    cmdMortgage.Visible = False
    cmdBankrupt.Visible = False
    cmdBuild.Visible = False
End Sub

'############### waiting window - Begin
Public Sub viewWaitingWin()
    lblWGameNo.Caption = tableID
    lblWGameTitle.Caption = gameTitle
    lblWPID.Caption = PlayerID
    lblWPName.Caption = PlayerName
    optRules(currentRules).value = True
    If player(PlayerNumber).keyPlayer Then
        cmdReady.Visible = False
        cmdStart.Visible = True
        optRules(0).Enabled = True
        optRules(1).Enabled = True
    Else
        cmdReady.Visible = True
        cmdStart.Visible = False
        optRules(0).Enabled = False
        optRules(1).Enabled = False
    End If
    Call loadToken
    imgWToken_Click (player(PlayerNumber).tokenID)
    picWPInfoBox(1).BackColor = GetSetting("Monopoly", "Game Setting", "Player Slot Color 1", "&H8080FF")
    picWPInfoBox(2).BackColor = GetSetting("Monopoly", "Game Setting", "Player Slot Color 2", "&HFF8080")
    picWPInfoBox(3).BackColor = GetSetting("Monopoly", "Game Setting", "Player Slot Color 3", "&H80FF80")
    picWPInfoBox(4).BackColor = GetSetting("Monopoly", "Game Setting", "Player Slot Color 4", "&H80FFFF")

    For loopX = 1 To 4
        lblWPJoinName(loopX).Caption = ""
        lblWReady(loopX).Caption = ""
        imgWPToken(loopX).Picture = LoadPicture("")
        If player(loopX).tokenID <> 0 Then
            imgWPToken(loopX).Picture = LoadPicture(App.Path & "/images/token/" & token(player(loopX).tokenID).file)
        End If
        If player(loopX).PID <> 0 Then
            lblWPJoinName(loopX).Caption = player(loopX).PName
        End If
        If player(loopX).ready Then
            lblWReady(loopX).Caption = "Ready"
        End If
    Next
    picWaiting.Visible = True
End Sub

Public Sub loadToken()
    Dim intTop As Long
    Dim intLeft As Long
    intTop = 1800
    intLeft = 240
    For loopX = 1 To 10
        Load imgWToken(loopX)
        imgWToken(loopX).Picture = LoadPicture("images/Token/" & token(loopX).file)
        imgWToken(loopX).Visible = True
        imgWToken(loopX).Top = intTop
        imgWToken(loopX).Left = intLeft
        intLeft = intLeft + 720
        If loopX = 5 Then
            intTop = intTop + 720
            intLeft = 240
        End If
    Next
    imgWToken(0).Visible = False
End Sub

Private Sub imgWToken_Click(Index As Integer)
    If Index >= 1 And Index <= 10 Then
        If player(PlayerNumber).tokenID <> 0 Then
            imgWToken(player(PlayerNumber).tokenID).BorderStyle = 0
        End If
        imgWToken(Index).BorderStyle = 1
    End If
    frmMain.wskClient.SendData "cmd;changeToken;" & tableID & ";" & PlayerNumber & ";" & Index & ";"
    DoEvents
End Sub

Public Sub changeToken(number As Integer, tokenNumber As Integer)
    If tokenNumber > 0 Then
        player(number).tokenID = tokenNumber
        imgWPToken(number).Picture = LoadPicture(App.Path & "/images/Token/" & token(player(number).tokenID).file)
    End If
End Sub

Private Sub optRules_Click(Index As Integer)
    frmMain.wskClient.SendData "cmd;changeRules;" & tableID & ";" & PlayerNumber & ";" & Index & ";"
    DoEvents
End Sub

Public Sub changeRules(number As Integer, rulesNumber As Integer)
    currentRules = rulesNumber
    optRules(rulesNumber).value = True
End Sub

Private Sub cmdReady_Click()
    frmMain.wskClient.SendData "cmd;playerReady;" & tableID & ";" & PlayerNumber & ";"
    DoEvents
End Sub

Public Sub playerReady(number As Integer)
    player(number).ready = True
    lblWReady(number).Caption = "Ready"
End Sub

Private Sub cmdStart_Click()
    Dim allReady As Boolean
    Dim pCount As Integer

    If player(PlayerNumber).keyPlayer Then
        allReady = True
        pCount = 0
        For z = 1 To maxPlayer
            If player(z).PID <> 0 Then
                pCount = pCount + 1
            End If
            If z <> PlayerNumber Then
                If player(z).PID <> 0 And Not player(z).ready Then
                    allReady = False
                    Exit For
                End If
            End If
        Next
        If allReady And pCount >= 2 And player(PlayerNumber).keyPlayer Then
            Dim strInit As String
            strInit = ""
            For i = 1 To 16
                rndChance(i) = i
                rndCommunity(i) = i
            Next
            For i = 1 To 20
                Randomize
                num1 = (Rnd() * 100 Mod 15) + 1
                Randomize
                num2 = (Rnd() * 100 Mod 15) + 1
                If num1 <> num2 Then
                    temp = rndChance(num1)
                    rndChance(num1) = rndChance(num2)
                    rndChance(num2) = temp
                End If
                Randomize
                num1 = (Rnd() * 100 Mod 15) + 1
                Randomize
                num2 = (Rnd() * 100 Mod 15) + 1
                If num1 <> num2 Then
                    temp = rndCommunity(num1)
                    rndCommunity(num1) = rndCommunity(num2)
                    rndCommunity(num2) = temp
                End If
            Next
            For i = 1 To 16
                strInit = strInit & rndChance(i) & "," & rndCommunity(i) & ";"
            Next
            
            If gameRules(currentRules).startingProperties > 0 Then
                For z = 1 To 4
                    If player(z).PID <> 0 Then
                        numProperties = 0
                        Do
                            Randomize
                            temp = (Rnd() * 100 Mod 40) + 1
                            If deed(temp).mortgageValue > 0 And Not slot(temp).hasOwner Then
                                slot(temp).hasOwner = True
                                slot(temp).ownerPos = z
                                strInit = strInit & z & "," & temp & ";"
                                numProperties = numProperties + 1
                            End If
                        Loop Until numProperties = gameRules(currentRules).startingProperties
                    End If
                Next
            End If
            frmMain.wskClient.SendData "table;startGame;" & tableID & ";" & strInit & "EOT;"
            DoEvents
        End If
    End If
End Sub

Private Sub cmdWExit_Click()
    Unload Me
End Sub
'############### waiting window - End

'############### Init game  - Begin
Public Sub initGameVar(str As String)
    Dim c As Integer
    currentPlayer = 1
    
    Call drawMonopolyBoard
    
    For loopX = 1 To 4
        If player(loopX).PID <> 0 Then
            player(loopX).currentSlot = 1
            player(loopX).inParking = False
            player(loopX).inJail = False
            player(loopX).numTurnInJail = 0
            player(loopX).cardOutOfJailA = False
            player(loopX).cardOutOfJailB = False
            player(loopX).cash = gameRules(currentRules).initialCash
        End If
    Next

    For loopX = 1 To 4
        If player(loopX).PID <> 0 Then
            Call setPlayerStatus("pname", loopX, player(loopX).PName)
            Call setPlayerStatus("cash", loopX, player(loopX).cash)
            If player(loopX).tokenID <> 0 Then
                imgPlayerToken(loopX).Picture = LoadPicture("images/Token/" & token(player(loopX).tokenID).file)
            End If
        End If
    Next
    housesAvailable = gameRules(currentRules).totalHouses
    hotelsAvailable = gameRules(currentRules).totalHotels
    
    currentCommunityCard = 1
    currentChanceCard = 1
    For c = 1 To 16
        rndChance(c) = Int(Split(Split(str, ";")(c), ",")(0))
        rndCommunity(c) = Int(Split(Split(str, ";")(c), ",")(1))
    Next
    c = 17
    Do
        If Split(str, ";")(c) <> "EOT" Then
            Call buyProperties(Int(Split(Split(str, ";")(c), ",")(0)), Int(Split(Split(str, ";")(c), ",")(1)))
        End If
        c = c + 1
    Loop Until Split(str, ";")(c - 1) = "EOT"
    Call StartGame
End Sub

Public Function drawMonopolyBoard()
    Dim tmpImgTop As Long
    Dim tmpImgLeft As Long
    On Error Resume Next
    tmpImgTop = startTop + slotSize + (slotSizeW * 9)
    tmpImgLeft = startLeft + slotSize + (slotSizeW * 9)
    For p = 1 To 40
        Load imgPlayerDeed1(p)
        Load imgPlayerDeed2(p)
        Load imgPlayerDeed3(p)
        Load imgPlayerDeed4(p)
        Load imgMortgageDeedCard(p)
        Load imgTradeCard1(p)
        Load imgTradeCard2(p)
        Load imgTradingCard1(p)
        Load imgTradingCard2(p)
        Load imgSlot(p)
        Load shpSlot(p)
        Load imgHouse1(p)
        Load imgHouse3(p)
        Load imgHouse2(p)
        Load imgHouse4(p)
        Load imgHotel(p)
        
        slot(p).deedID = p
        slot(p).onMortgage = False
        slot(p).hasOwner = False
        slot(p).ownerPos = 0
        slot(p).tokenSlot(1) = False
        slot(p).tokenSlot(2) = False
        slot(p).tokenSlot(3) = False
        slot(p).tokenSlot(4) = False
        
        imgSlot(p).Height = slotSize
        imgSlot(p).Width = slotSize
        imgSlot(p).Top = tmpImgTop
        imgSlot(p).Left = tmpImgLeft
        imgSlot(p).Visible = False
        imgSlot(p).Stretch = True
        imgSlot(p).Picture = LoadPicture(App.Path & "\images\slot\slot" & p & ".jpg")
        lblLoading.Caption = "\images\slot\slot" & p & ".jpg"
        
        shpSlot(p).Height = slotSize
        shpSlot(p).Width = slotSize
        shpSlot(p).Top = tmpImgTop
        shpSlot(p).Left = tmpImgLeft
        shpSlot(p).Visible = False
        shpSlot(p).ZOrder
        
        imgHouse1(p).Height = slotSizeW / 4
        imgHouse1(p).Width = slotSizeW / 4
        imgHouse1(p).Visible = False
        imgHouse1(p).ZOrder
 
        imgHouse2(p).Height = slotSizeW / 4
        imgHouse2(p).Width = slotSizeW / 4
        imgHouse2(p).Visible = False
        imgHouse2(p).ZOrder

        imgHouse3(p).Height = slotSizeW / 4
        imgHouse3(p).Width = slotSizeW / 4
        imgHouse3(p).Visible = False
        imgHouse3(p).ZOrder

        imgHouse4(p).Height = slotSizeW / 4
        imgHouse4(p).Width = slotSizeW / 4
        imgHouse4(p).Visible = False
        imgHouse4(p).ZOrder

        imgHotel(p).Height = slotSizeW / 2
        imgHotel(p).Width = slotSizeW / 2
        imgHotel(p).Visible = False
        imgHotel(p).ZOrder
        
        imgHouse1(p).Picture = LoadPicture("images/house.gif")
        imgHouse2(p).Picture = LoadPicture("images/house.gif")
        imgHouse3(p).Picture = LoadPicture("images/house.gif")
        imgHouse4(p).Picture = LoadPicture("images/house.gif")
        imgHotel(p).Picture = LoadPicture("images/hotel.gif")
        
        Select Case p
            Case 1 To 10
                imgHouse1(p).Top = tmpImgTop + slotSize - slotSizeW / 4
                imgHouse1(p).Left = tmpImgLeft + slotSizeW - slotSizeW / 4
                imgHouse2(p).Top = tmpImgTop + slotSize - slotSizeW / 4
                imgHouse2(p).Left = tmpImgLeft + slotSizeW / 2
                imgHouse3(p).Top = tmpImgTop + slotSize - slotSizeW / 4
                imgHouse3(p).Left = tmpImgLeft + slotSizeW / 4
                imgHouse4(p).Top = tmpImgTop + slotSize - slotSizeW / 4
                imgHouse4(p).Left = tmpImgLeft
                imgHotel(p).Top = tmpImgTop + slotSize - slotSizeW / 2
                imgHotel(p).Left = tmpImgLeft + slotSizeW / 4
            Case 11 To 20
                imgHouse1(p).Top = tmpImgTop + slotSizeW - slotSizeW / 4
                imgHouse1(p).Left = tmpImgLeft
                imgHouse2(p).Top = tmpImgTop + slotSizeW / 2
                imgHouse2(p).Left = tmpImgLeft
                imgHouse3(p).Top = tmpImgTop + slotSizeW / 4
                imgHouse3(p).Left = tmpImgLeft
                imgHouse4(p).Top = tmpImgTop
                imgHouse4(p).Left = tmpImgLeft
                imgHotel(p).Top = tmpImgTop + slotSizeW / 4
                imgHotel(p).Left = tmpImgLeft
            Case 21 To 30
                Load imgHouse1(p)
                imgHouse1(p).Top = tmpImgTop
                imgHouse1(p).Left = tmpImgLeft
                imgHouse2(p).Top = tmpImgTop
                imgHouse2(p).Left = tmpImgLeft + slotSizeW / 4
                imgHouse3(p).Top = tmpImgTop
                imgHouse3(p).Left = tmpImgLeft + slotSizeW / 2
                imgHouse4(p).Top = tmpImgTop
                imgHouse4(p).Left = tmpImgLeft + slotSizeW - slotSizeW / 4
                imgHotel(p).Top = tmpImgTop
                imgHotel(p).Left = tmpImgLeft + slotSizeW / 4
            Case 31 To 40
                imgHouse1(p).Top = tmpImgTop
                imgHouse1(p).Left = tmpImgLeft + slotSize - slotSizeW / 4
                imgHouse2(p).Top = tmpImgTop + slotSizeW / 4
                imgHouse2(p).Left = tmpImgLeft + slotSize - slotSizeW / 4
                imgHouse3(p).Top = tmpImgTop + slotSizeW / 2
                imgHouse3(p).Left = tmpImgLeft + slotSize - slotSizeW / 4
                imgHouse4(p).Top = tmpImgTop + slotSizeW - slotSizeW / 4
                imgHouse4(p).Left = tmpImgLeft + slotSize - slotSizeW / 4
                imgHotel(p).Top = tmpImgTop + slotSizeW / 4
                imgHotel(p).Left = tmpImgLeft + slotSize - slotSizeW / 2
        End Select
                
        If (p Mod 10) = 1 Then
            Select Case ((-Int(-(p / 10))) Mod 4)
                Case 1: tmpImgLeft = tmpImgLeft - slotSizeW
                Case 2: tmpImgTop = tmpImgTop - slotSizeW
                        tmpImgLeft = tmpImgLeft + slotSizeW - slotSize
                        imgSlot(p).Left = tmpImgLeft
                        shpSlot(p).Left = tmpImgLeft
                Case 3: tmpImgLeft = tmpImgLeft + slotSize
                        tmpImgTop = tmpImgTop + slotSizeW - slotSize
                        imgSlot(p).Top = tmpImgTop
                        shpSlot(p).Top = tmpImgTop
                Case 0: tmpImgTop = tmpImgTop + slotSize
            End Select
        Else
            Select Case ((-Int(-(p / 10))) Mod 4)
                Case 1
                    imgSlot(p).Width = slotSizeW
                    imgSlot(p).Left = tmpImgLeft
                    shpSlot(p).Width = slotSizeW
                    shpSlot(p).Left = tmpImgLeft
                    shpSlot(p).Top = shpSlot(p).Top + slotSize - 50
                    shpSlot(p).Height = 100
                    tmpImgLeft = tmpImgLeft - slotSizeW
                Case 2
                    imgSlot(p).Height = slotSizeW
                    imgSlot(p).Top = tmpImgTop
                    shpSlot(p).Height = slotSizeW
                    shpSlot(p).Top = tmpImgTop
                    shpSlot(p).Left = shpSlot(p).Left - 50
                    shpSlot(p).Width = 100
                    tmpImgTop = tmpImgTop - slotSizeW
                Case 3
                    imgSlot(p).Width = slotSizeW
                    imgSlot(p).Left = tmpImgLeft
                    shpSlot(p).Width = slotSizeW
                    shpSlot(p).Left = tmpImgLeft
                    shpSlot(p).Top = shpSlot(p).Top - 50
                    shpSlot(p).Height = 100
                    tmpImgLeft = tmpImgLeft + slotSizeW
                Case 0
                    imgSlot(p).Height = slotSizeW
                    imgSlot(p).Top = tmpImgTop
                    shpSlot(p).Height = slotSizeW
                    shpSlot(p).Top = tmpImgTop
                    shpSlot(p).Left = shpSlot(p).Left + slotSize - 50
                    shpSlot(p).Width = 100
                    tmpImgTop = tmpImgTop + slotSizeW
            End Select
        End If
    Next
    
    picWaiting.Visible = False
    For p = 1 To 40
        imgSlot(p).Visible = True
    Next
End Function

Public Sub StartGame()
    Dim temp As Integer
    Dim startProperties
    tempProperties = ""
    For z = 1 To 4
        If player(z).PID <> 0 Then
            If player(z).tokenID <> 0 Then
                imgJail(z).Visible = player(z).inJail
                picToken(z).BackColor = picPlayer(z).BackColor
                picToken(z).Height = slotSizeW / 3
                picToken(z).Width = slotSizeW / 3
                picToken(z).Top = imgSlot(player(z).currentSlot).Top + 100
                picToken(z).Left = imgSlot(player(z).currentSlot).Left + (50 + picToken(z).Width) * (z - 1)
                picToken(z).Visible = True
                imgToken(z).Picture = LoadPicture("images/Token/" & token(player(z).tokenID).file)
                imgToken(z).Height = slotSizeW / 3
                imgToken(z).Width = slotSizeW / 3
            End If
        End If
    Next
    
    picSideBar.Visible = True
    picChatBox.Visible = True
    cmdTrade.Visible = True
    cmdTrade.Enabled = True
    cmdStatus.Visible = True
    cmdStatus.Enabled = True
    gameStarted = True

    lblGameID.Caption = "ID: " & tableID
    lblGameTitle.Caption = "Title: " & gameTitle
    Call nextPlayer
End Sub
'############### Init game  - End

'############### In game function  - Begin
Private Sub cmdBankrupt_Click()
    If MsgBox("Are your sure?", vbQuestion + vbYesNo, "Bankrupt") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdCloseStatus_Click()
    picStatus.Visible = False
    frmButton.Enabled = True
End Sub

Private Sub cmdMenu_Click()
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit Game") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdStatus_Click()
    picStatus.Visible = True
    picStatus.ZOrder
    frmButton.Enabled = False
    cmdPlayerStatus.value = True
End Sub

Private Sub cmdPlayerStatus_Click()
    Dim img As Image
    Dim tmpLeft As Long
    Dim tmpTop As Long
    Dim deedCount As Integer
    Dim j As Integer
    For z = 1 To 4
        deedCount = 0
        tmpTop = 400
        tmpLeft = 100
        If player(z).PID <> 0 Then
            lblStatusPlayerName(z).Caption = player(z).PName & " [$" & player(z).cash & "]"
            lblStatusPlayerName(z).BackColor = picPlayer(z).BackColor
            picStatusDeed(z).Visible = True
            Dim assets As Currency
            assets = 0
            For j = 1 To 40
                If slot(j).hasOwner And slot(j).ownerPos = z Then
                    totalWorth = totalWorth + deed(j).mortgageValue
                    If slot(j).numOfHouses >= 0 And slot(j).numOfHouses < gameRules(currentRules).housesPerHotel Then
                        assets = assets + deed(j).mortgageValue + deed(j).houseCost * slot(j).numOfHouses
                    Else
                        assets = assets + deed(j).mortgageValue + deed(j).houseCost * (gameRules(currentRules).housesPerHotel - 1) + deed(j).hotelCost
                    End If
                End If
            Next
            lblAssets(z).Caption = assets
            lblNW(z).Caption = assets + player(z).cash
            For i = 1 To 40
                Select Case z
                    Case 1: Set img = imgPlayerDeed1(i)
                    Case 2: Set img = imgPlayerDeed2(i)
                    Case 3: Set img = imgPlayerDeed3(i)
                    Case 4: Set img = imgPlayerDeed4(i)
                End Select
                If deed(i).mortgageValue > 0 Then
                    If slot(i).hasOwner And slot(i).ownerPos = z Then
                        If Not slot(i).onMortgage Then
                            img.Picture = LoadPicture("images/deed/deed" & i & ".jpg")
                        Else
                            img.Picture = LoadPicture("images/deed/deedM" & i & ".jpg")
                        End If
                        img.Visible = True
                        img.ZOrder
                    Else
                        img.Visible = False
                    End If
                    img.Top = tmpTop
                    img.Left = tmpLeft
                    tmpTop = tmpTop + 325
                    deedCount = deedCount + 1
                End If
                If deedCount > 13 Then
                    tmpTop = 400
                    tmpLeft = tmpLeft + 1200
                    deedCount = 0
                End If
            Next
        Else
            picStatusDeed(z).Visible = False
        End If
    Next
End Sub

Public Sub nextPlayer()
    cmdBuy.Visible = False
    cmdSell.Visible = False
    cmdDone.Visible = False
    cmdPay.Visible = False
    cmdPayTax.Visible = False
    cmdPayPercentTax.Visible = False
    cmdUseCard.Visible = False
    cmdAuction.Visible = False
    cmdBankrupt.Visible = False
    cmdBuild.Visible = False
    cmdMortgage.Visible = True
    cmdRollDice.Visible = True
    If PlayerNumber = currentPlayer Then
        cmdMortgage.Enabled = True
        cmdRollDice.Enabled = True
        If player(currentPlayer).inJail Then
            cmdPay.Visible = True
            cmdPay.Enabled = True
            If player(currentPlayer).cardOutOfJailA Or player(currentPlayer).cardOutOfJailB Then
                cmdUseCard.Visible = True
                cmdUseCard.Enabled = True
            End If
        Else
            cmdPay.Visible = False
            cmdPay.Enabled = False
            cmdUseCard.Visible = False
            cmdUseCard.Enabled = False
        End If
    Else
        cmdMortgage.Enabled = False
        cmdRollDice.Enabled = False
        If player(currentPlayer).inJail Then
            cmdPay.Visible = True
            cmdPay.Enabled = False
            If player(currentPlayer).cardOutOfJailA Or player(currentPlayer).cardOutOfJailB Then
                cmdUseCard.Visible = True
                cmdUseCard.Enabled = False
            End If
        Else
            cmdPay.Visible = False
            cmdPay.Enabled = False
            cmdUseCard.Visible = False
            cmdUseCard.Enabled = False
        End If
    End If
    
    If player(currentPlayer).inJail Then
        player(currentPlayer).numTurnInJail = player(currentPlayer).numTurnInJail + 1
        If player(currentPlayer).numTurnInJail = 3 Then
            cmdRollDice.Enabled = False
        End If
    End If
        
    For z = 1 To maxPlayer
        If z = currentPlayer Then
            picPlayer(z).BorderStyle = 1
        Else
            picPlayer(z).BorderStyle = 0
        End If
    Next
    
    imgCard.Picture = LoadPicture("")
    imgCurPToken.Picture = LoadPicture("images/token/" & token(player(currentPlayer).tokenID).file)
    If player(currentPlayer).inJail Then
        imgCurJail.Visible = True
    Else
        imgCurJail.Visible = False
    End If
    lblCurPName.Caption = player(currentPlayer).PName
    lblCurCash.Caption = "$" & player(currentPlayer).cash
End Sub

Private Sub cmdPayTax_Click()
    frmMain.wskClient.SendData "cmd;payTax;" & tableID & ";" & PlayerNumber & ";" & player(PlayerNumber).currentSlot & ";"
    DoEvents
End Sub

Public Sub payTax(number As Integer)
    Call setPlayerStatus("cash", number, player(number).cash - 200)

    cmdPayTax.Enabled = False
    cmdPayPercentTax.Enabled = False
    cmdDone.Enabled = False

    If number = PlayerNumber Then
        cmdDone.Enabled = True
    End If
    loadPlayerInfo (currentPlayer)
End Sub

Private Sub cmdPayPercentTax_Click()
    frmMain.wskClient.SendData "cmd;payPercentTax;" & tableID & ";" & PlayerNumber & ";" & player(PlayerNumber).currentSlot & ";"
    DoEvents
End Sub

Public Sub payPercentTax(number As Integer)
    Dim totalWorth As Currency
    totalWorth = player(number).cash
    For z = 1 To 40
        If slot(z).hasOwner And slot(z).ownerPos = number Then
            totalWorth = totalWorth + deed(z).mortgageValue
            If slot(z).numOfHouses >= 0 And slot(z).numOfHouses < gameRules(currentRules).housesPerHotel Then
                totalWorth = totalWorth + deed(z).mortgageValue + deed(z).houseCost * slot(z).numOfHouses
            Else
                totalWorth = totalWorth + deed(z).mortgageValue + deed(z).houseCost * (gameRules(currentRules).housesPerHotel - 1) + deed(z).hotelCost
            End If
        End If
    Next
    Call setPlayerStatus("cash", number, player(number).cash - Int(totalWorth * 10 / 100))

    cmdPayTax.Enabled = False
    cmdPayPercentTax.Enabled = False
    cmdDone.Enabled = False

    If number = PlayerNumber Then
        cmdDone.Enabled = True
    End If
    loadPlayerInfo (currentPlayer)
End Sub

Private Sub cmdPay_Click()
    frmMain.wskClient.SendData "cmd;payFine;" & tableID & ";" & PlayerNumber & ";"
    DoEvents
End Sub

Public Sub payJailFine(number As Integer)
    Call setPlayerStatus("cash", number, player(number).cash - 50)
    player(number).inJail = False
    player(number).numTurnInJail = 0
    imgJail(number).Visible = player(number).inJail
    cmdPay.Visible = False
    cmdUseCard.value = False
    cmdRollDice.Visible = True
    If PlayerNumber = number Then
        cmdRollDice.Enabled = True
    Else
        cmdRollDice.Enabled = False
    End If
End Sub

Private Sub cmdUseCard_Click()
    frmMain.wskClient.SendData "cmd;useCard;" & tableID & ";" & PlayerNumber & ";"
    DoEvents
End Sub

Public Sub useCard(number As Integer)
    cmdUseCard.Enabled = False
    cmdUseCard.Visible = False
    If player(number).cardOutOfJailA Or player(number).cardOutOfJailB Then
        If player(number).cardOutOfJailA Then
            player(number).cardOutOfJailA = False
        Else
            player(number).cardOutOfJailB = False
        End If
        player(number).inJail = False
        player(number).numTurnInJail = 0
        imgJail(number).Visible = player(number).inJail
        cmdPay.Visible = False
        cmdUseCard.value = False
        cmdRollDice.Visible = True
        If PlayerNumber = number Then
            cmdRollDice.Enabled = True
        Else
            cmdRollDice.Enabled = False
        End If
    End If
End Sub

Private Sub cmdRollDice_Click()
    Randomize
    
    tempDice = (Rnd() * 100 Mod 12) + 6
    tmrDice.Enabled = True
    cmdRollDice.Enabled = False
    cmdPay.Enabled = False
    cmdPay.Visible = False
    cmdUseCard.Enabled = False
    cmdUseCard.value = False
End Sub

Public Function rollDice() As Integer
    Randomize
    rollDice = Int(Rnd * 6 + 1)
End Function

Private Sub tmrDice_Timer()
    numDice1 = rollDice
    numDice2 = rollDice
    s = play("dice.wav")
    imgDice1.Picture = LoadPicture(App.Path & "/images/dice/num" & numDice1 & ".gif")
    imgDice2.Picture = LoadPicture(App.Path & "/images/dice/num" & numDice2 & ".gif")
    tempDice = tempDice - 1
    If tempDice <= 0 Then
        tmrDice.Enabled = False
        frmMain.wskClient.SendData "cmd;rollDiceResult;" & tableID & ";" & PlayerNumber & ";" & numDice1 & ";" & numDice2 & ";"
        DoEvents
    End If
End Sub

Public Function moveToken(number As Integer, numOfMove As Integer)
    tempNumber = number
    tempNumOfMove = numOfMove
    tempMoveDest = player(number).currentSlot + numOfMove
    If tempMoveDest > 40 Then
        tempMoveDest = tempMoveDest - 40
    End If
    tempMoveCheck = True
    tmrMoveToken.Enabled = True
 End Function

Public Function moveTokenTo(number As Integer, moveToSlot As Integer)
    tempNumber = number
    tempMoveDest = moveToSlot
    tempMoveCheck = False
    tmrMoveToken.Enabled = True
End Function

Private Sub tmrMoveToken_Timer()
    If tempMoveDest - player(tempNumber).currentSlot >= 0 Then
        player(tempNumber).currentSlot = player(tempNumber).currentSlot + 1
    ElseIf tempMoveDest - player(tempNumber).currentSlot < 0 And tempMoveDest - player(tempNumber).currentSlot >= -10 Then
        player(tempNumber).currentSlot = player(tempNumber).currentSlot - 1
    Else
        player(tempNumber).currentSlot = player(tempNumber).currentSlot + 1
    End If
    If player(tempNumber).currentSlot > 40 Then
        If Not player(tempNumber).inJail Then
            Call setPlayerStatus("cash", tempNumber, player(tempNumber).cash + gameRules(currentRules).salary)
        End If
        player(tempNumber).currentSlot = player(tempNumber).currentSlot - 40
    End If
    If tempNumber = 1 Or tempNumber = 2 Then
        picToken(tempNumber).Move imgSlot(player(tempNumber).currentSlot).Left + 20 + (50 + picToken(tempNumber).Width) * (tempNumber - 1), imgSlot(player(tempNumber).currentSlot).Top + 100
    Else
        picToken(tempNumber).Move imgSlot(player(tempNumber).currentSlot).Left + 20 + (50 + picToken(tempNumber).Width) * (tempNumber - 3), imgSlot(player(tempNumber).currentSlot).Top + 100 + picToken(tempNumber).Height
    End If
    s = play("token.wav")
    If player(tempNumber).currentSlot = tempMoveDest Then
        tmrMoveToken.Enabled = False
        If tempMoveCheck Then
            Call playSlotSound(tempMoveDest)
            Call checkMove(tempNumber, tempMoveDest)
            Call checkMove2(tempNumber, tempMoveDest)
        End If
    End If
End Sub

Public Sub playSlotSound(dest As Integer)
    If deed(dest).soundFile <> "" Then
        s = play(deed(dest).soundFile)
    End If
End Sub

Public Sub checkMove(number As Integer, dest As Integer)
    Select Case dest
        Case 1, 11
                '...
        Case 21
                player(number).inParking = True
        Case 31 'go to jail
                Call moveTokenTo(number, 11)
                player(number).inJail = True
                player(number).numTurnInJail = 0
                imgJail(number).Visible = player(number).inJail
        Case 3, 18, 34 'community chest
                community (number)
        Case 8, 23, 37 'chance
                chance (number)
        Case 5 'income tax
                cmdPayPercentTax.Visible = True
                cmdPayTax.Visible = True
                If number = PlayerNumber Then
                    cmdPayPercentTax.Enabled = True
                    cmdPayTax.Enabled = True
                Else
                    cmdPayPercentTax.Enabled = False
                    cmdPayTax.Enabled = False
                End If
        Case 39 'luxury tax
                Call setPlayerStatus("cash", number, player(number).cash - 75)
        Case 13, 29 'utilities
                If slot(dest).hasOwner Then
                    Dim utilitiesRental As Currency
                    If slot(dest).hasOwner Then
                        If slot(13).ownerPos = slot(29).ownerPos Then
                            utilitiesRental = ((numDice1 + numDice2) * 10)
                        Else
                            utilitiesRental = ((numDice1 + numDice2) * 4)
                        End If
                        Call setPlayerStatus("cash", number, player(number).cash - utilitiesRental)
                        Call setPlayerStatus("cash", slot(dest).ownerPos, player(slot(dest).ownerPos).cash + utilitiesRental)
                    End If
                End If
        Case 6, 16, 26, 36 'railroads
                If slot(dest).hasOwner Then
                    Dim NumOfRR As Integer
                    Dim TotalRental As Currency
                    If slot(dest).hasOwner Then
                        NumOfRR = 0
                        For z = 6 To 36 Step 10
                            If slot(z).ownerPos = slot(dest).ownerPos Then
                                NumOfRR = NumOfRR + 1
                            End If
                        Next
                        TotalRental = deed(dest).rentHouse(NumOfRR - 1)
                        Call setPlayerStatus("cash", number, player(number).cash - TotalRental)
                        Call setPlayerStatus("cash", slot(dest).ownerPos, player(slot(dest).ownerPos).cash + TotalRental)
                    End If
                End If
        Case Else
                If slot(dest).hasOwner Then
                    If Not player(slot(dest).ownerPos).inParking Then
                        Call setPlayerStatus("cash", number, player(number).cash - deed(dest).rentHouse(slot(dest).numOfHouses))
                        Call setPlayerStatus("cash", slot(dest).ownerPos, player(slot(dest).ownerPos).cash + deed(dest).rentHouse(slot(dest).numOfHouses))
                    End If
                End If
    End Select
    Call checkMove2(number, dest)
End Sub

Public Sub checkMove2(number As Integer, dest As Integer)
    If dest <> 5 Then
        If number = PlayerNumber Then
            If Not slot(dest).hasOwner And deed(dest).mortgageValue > 0 Then
                If numDice1 <> numDice2 Then
                    cmdBuy.Enabled = True
                    cmdAuction.Enabled = True
                Else
                    cmdDone.Enabled = True
                End If
            Else
                cmdDone.Enabled = True
            End If
        End If
    End If
    
    If player(number).cash < 0 And number = PlayerNumber Then
        cmdDone.Enabled = False
        cmdMortgage.Visible = True
        cmdMortgage.Enabled = True
        cmdBankrupt.Visible = True
        cmdBankrupt.Enabled = True
    End If
End Sub

Public Sub rollDiceResult(number As Integer, dice1 As Integer, dice2 As Integer)
    cmdRollDice.Visible = False
    
    cmdBuy.Visible = True
    cmdAuction.Visible = True
    cmdDone.Visible = True
    
    cmdBuy.Enabled = False
    cmdAuction.Enabled = False
    cmdDone.Enabled = False
    
    numDice1 = dice1
    numDice2 = dice2
    
    imgDice1.Picture = LoadPicture(App.Path & "/images/dice/num" & dice1 & ".gif")
    imgDice2.Picture = LoadPicture(App.Path & "/images/dice/num" & dice2 & ".gif")
    
    If dice1 = dice2 Then
        If Not player(number).inJail Then
            rollDiceDblCount = rollDiceDblCount + 1
            If rollDiceDblCount = 3 Then
                player(number).inJail = True
                player(number).numTurnInJail = 0
                imgJail(number).Visible = player(number).inJail
                Call moveTokenTo(number, 11)
            End If
        Else
            player(number).inJail = False
            player(number).numTurnInJail = 0
            imgJail(number).Visible = player(number).inJail
            numDice1 = numDice1 + 1
            numDice2 = numDice2 - 1
        End If
    End If

    If Not player(number).inJail Then
        Call moveToken(number, dice1 + dice2)
    End If
End Sub

Private Sub cmdDone_Click()
    frmMain.wskClient.SendData "cmd;done;" & tableID & ";" & PlayerNumber & ";"
    DoEvents
End Sub

Public Sub done(number As Integer)
    If (numDice1 <> numDice2) Or (player(number).inJail And player(number).numTurnInJail = 0) Then
        For z = 1 To maxPlayer
            currentPlayer = currentPlayer + 1
            If currentPlayer > maxPlayer Then currentPlayer = 1
            If player(currentPlayer).PID <> 0 Then Exit For
        Next
        rollDiceDblCount = 0
    End If
    loadPlayerInfo (currentPlayer)
    Call nextPlayer
End Sub

Private Sub cmdTrade_Click()
    Dim tempbtnLeft As Integer
    Dim tempbtntop As Integer
    picSelectTrader.Visible = True
    picSelectTrader.ZOrder
    tempbtnLeft = 840
    tempbtntop = 600
    For z = 1 To maxPlayer
        If player(z).PID <> 0 And z <> PlayerNumber Then
            cmdPlayerName(z).Left = tempbtnLeft
            cmdPlayerName(z).Top = tempbtntop
            cmdPlayerName(z).Visible = True
            cmdPlayerName(z).BackColor = picPlayer(z).BackColor
            cmdPlayerName(z).Caption = player(z).PName
            tempbtntop = tempbtntop + 600
        Else
            cmdPlayerName(z).Visible = False
        End If
    Next
End Sub

Private Sub cmdCancelTrading_Click()
    frmButton.Enabled = True
    picSelectTrader.Visible = False
End Sub

Private Sub cmdPlayerName_Click(Index As Integer)
    tradingWith = Index
    picSelectTrader.Visible = False
    Call makeProposal
End Sub

Public Sub makeProposal()
    picMakeProposal.Visible = True
    picMakeProposal.ZOrder
    For i = 1 To 20
        propertiesTrade1(i) = 0
        propertiesTrade2(i) = 0
    Next
    txtTradeAmount1.Text = 0
    txtTradeAmount2.Text = 0
    frmButton.Enabled = False
    cmdProposeTrade.Visible = True
    cmdDoneTrade.Visible = False
    cmdCancelTrade.Visible = True
    cmdCounterTrade.Visible = False
    cmdRejectTrade.Visible = False
    cmdAcceptTrade.Visible = False
    Call loadTradingProperties(tradingWith)
    lblTradeAnounce.Caption = "Click propose to start trading with " & player(tradingWith).PName
End Sub

Public Sub loadTradingProperties(number As Integer)
    Dim tmpLeft As Long
    Dim tmpTop As Long
    Dim deedCount As Integer
    
    deedCount = 0
    tmpTop = 200
    tmpLeft = 200
    lblTraderName1.Caption = player(PlayerNumber).PName & " [$" & player(PlayerNumber).cash & "]"
    lblTraderName1.BackColor = picPlayer(PlayerNumber).BackColor
    lblTraderName2.Caption = player(number).PName & " [$" & player(number).cash & "]"
    lblTraderName2.BackColor = picPlayer(number).BackColor
    
    For i = 1 To 40
        If deed(i).mortgageValue > 0 Then
            If slot(i).hasOwner And slot(i).ownerPos = PlayerNumber Then
                If Not slot(i).onMortgage And slot(i).numOfHotels = 0 And slot(i).numOfHouses = 0 Then
                    imgTradeCard1(i).Picture = LoadPicture("images/deed/deed" & i & ".jpg")
                    imgTradeCard1(i).Visible = True
                Else
                    imgTradeCard1(i).Visible = False
                End If
                imgTradeCard1(i).Top = tmpTop
                imgTradeCard1(i).Left = tmpLeft
                imgTradeCard1(i).ZOrder
            Else
                imgTradeCard1(i).Visible = False
            End If
            If slot(i).hasOwner And slot(i).ownerPos = number Then
                If Not slot(i).onMortgage And slot(i).numOfHotels = 0 And slot(i).numOfHouses = 0 Then
                    imgTradeCard2(i).Picture = LoadPicture("images/deed/deed" & i & ".jpg")
                End If
                imgTradeCard2(i).Top = tmpTop
                imgTradeCard2(i).Left = tmpLeft
                imgTradeCard2(i).Visible = True
                imgTradeCard2(i).ZOrder
            Else
                imgTradeCard2(i).Visible = False
            End If
            tmpTop = tmpTop + 200
            deedCount = deedCount + 1
        End If
        If deedCount > 10 Then
            tmpTop = 200
            tmpLeft = tmpLeft + 960
            deedCount = 0
        End If
    Next
    If player(PlayerNumber).cardOutOfJailA Or player(PlayerNumber).cardOutOfJailB Then
        If player(PlayerNumber).cardOutOfJailA Then
            imgTradeOutJail1.Picture = LoadPicture("images/chance/chance1.jpg")
        Else
            imgTradeOutJail1.Picture = LoadPicture("images/community/community1.jpg")
        End If
        imgTradeOutJail1.Visible = True
    End If
    If player(number).cardOutOfJailA Or player(number).cardOutOfJailB Then
        If player(number).cardOutOfJailA Then
            imgTradeOutJail2.Picture = LoadPicture("images/chance/chance1.jpg")
        Else
            imgTradeOutJail2.Picture = LoadPicture("images/community/community1.jpg")
        End If
    End If
    deedCount = 0
    tmpTop = 200
    tmpLeft = 200
    For i = 1 To 20
        imgTradingCard1(i).Picture = LoadPicture("")
        imgTradingCard1(i).Visible = False
        imgTradingCard1(i).Top = tmpTop
        imgTradingCard1(i).Left = tmpLeft
        imgTradingCard1(i).ZOrder
        imgTradingCard2(i).Picture = LoadPicture("")
        imgTradingCard2(i).Visible = False
        imgTradingCard2(i).Top = tmpTop
        imgTradingCard2(i).Left = tmpLeft
        imgTradingCard2(i).ZOrder

        tmpTop = tmpTop + 200
        deedCount = deedCount + 1
        If deedCount > 4 Then
            tmpTop = 200
            tmpLeft = tmpLeft + 960
            deedCount = 0
        End If
    Next
    
    For z = 1 To 20
        If propertiesTrade1(z) <> 0 Then
            imgTradingCard1(z).Picture = LoadPicture("images/deed/deed" & propertiesTrade1(z) & ".jpg")
            imgTradingCard1(z).Visible = True
            imgTradeCard1(propertiesTrade1(z)).Visible = False
        End If
        If propertiesTrade2(z) <> 0 Then
            imgTradingCard2(z).Picture = LoadPicture("images/deed/deed" & propertiesTrade2(z) & ".jpg")
            imgTradingCard2(z).Visible = True
            imgTradeCard2(propertiesTrade2(z)).Visible = False
        End If
    Next
    If outOfJailTradeA Then
        Call imgTradeOutJail1_Click
    End If
    If outOfJailTradeB Then
        Call imgTradeOutJail2_Click
    End If
End Sub

Private Sub imgTradeOutJail1_Click()
    If player(PlayerNumber).cardOutOfJailA Or player(PlayerNumber).cardOutOfJailB Then
        If player(PlayerNumber).cardOutOfJailA Then
            imgTradingOutJail1.Picture = LoadPicture("images/chance/chance1.jpg")
        Else
            imgTradingOutJail1.Picture = LoadPicture("images/community/community1.jpg")
        End If
        outOfJailTradeA = True
        imgTradingOutJail1.Visible = True
        imgTradeOutJail1.Visible = False
    End If
End Sub

Private Sub imgTradingOutJail1_Click()
    imgTradeOutJail1.Visible = True
    imgTradingOutJail1.Visible = False
    outOfJailTradeA = False
End Sub

Private Sub imgTradeOutJail2_Click()
    If player(tradingWith).cardOutOfJailA Or player(tradingWith).cardOutOfJailB Then
        If player(tradingWith).cardOutOfJailA Then
            imgTradingOutJail2.Picture = LoadPicture("images/chance/chance1.jpg")
        Else
            imgTradingOutJail2.Picture = LoadPicture("images/community/community1.jpg")
        End If
        outOfJailTradeB = True
        imgTradingOutJail2.Visible = True
        imgTradeOutJail2.Visible = False
    End If
End Sub

Private Sub imgTradingOutJail2_Click()
    imgTradeOutJail2.Visible = True
    imgTradingOutJail2.Visible = False
    outOfJailTradeB = False
End Sub

Private Sub imgTradeCard1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imgTradeCardPreview.Picture = LoadPicture("images/deed/deed" & Index & ".jpg")
End Sub

Private Sub imgTradeCard2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imgTradeCardPreview.Picture = LoadPicture("images/deed/deed" & Index & ".jpg")
End Sub

Private Sub imgTradingCard1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If propertiesTrade1(Index) <> 0 Then
        imgTradeCardPreview.Picture = LoadPicture("images/deed/deed" & propertiesTrade1(Index) & ".jpg")
    End If
End Sub

Private Sub imgTradingCard2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If propertiesTrade2(Index) <> 0 Then
        imgTradeCardPreview.Picture = LoadPicture("images/deed/deed" & propertiesTrade2(Index) & ".jpg")
    End If
End Sub

Private Sub imgTradeCard1_Click(Index As Integer)
    For i = 1 To 20
        If propertiesTrade1(i) = 0 Then
            propertiesTrade1(i) = Index
            imgTradingCard1(i).Picture = LoadPicture("images/deed/deed" & Index & ".jpg")
            imgTradingCard1(i).Visible = True
            imgTradeCard1(Index).Visible = False
            Exit For
        End If
    Next
End Sub

Private Sub imgTradeCard2_Click(Index As Integer)
    For i = 1 To 20
        If propertiesTrade2(i) = 0 Then
            propertiesTrade2(i) = Index
            imgTradingCard2(i).Picture = LoadPicture("images/deed/deed" & Index & ".jpg")
            imgTradingCard2(i).Visible = True
            imgTradeCard2(Index).Visible = False
            Exit For
        End If
    Next
End Sub

Private Sub imgTradingCard1_Click(Index As Integer)
    imgTradeCard1(propertiesTrade1(Index)).Visible = True
    imgTradingCard1(Index).Visible = False
    propertiesTrade1(Index) = 0
End Sub

Private Sub imgTradingCard2_Click(Index As Integer)
    imgTradeCard2(propertiesTrade2(Index)).Visible = True
    imgTradingCard2(Index).Visible = False
    propertiesTrade2(Index) = 0
End Sub

Private Sub cmdProposeTrade_Click()
    Dim strTrade As String
    strTrade = "cmd;proposal;" & tableID & ";" & PlayerNumber
    strTrade = strTrade & ";" & tradingWith & ";" & outOfJailTradeA & ";" & outOfJailTradeB
    strTrade = strTrade & ";" & txtTradeAmount1 & ";" & txtTradeAmount2
    strTrade = strTrade & ";TP1"
    For z = 1 To 20
        If propertiesTrade1(z) <> 0 Then
            strTrade = strTrade & "," & propertiesTrade1(z)
        End If
    Next
    strTrade = strTrade & ",TP1-END;TPB"
    For z = 1 To 20
        If propertiesTrade2(z) <> 0 Then
            strTrade = strTrade & "," & propertiesTrade2(z)
        End If
    Next
    strTrade = strTrade & ",TP2-END;"
    frmMain.wskClient.SendData strTrade
    DoEvents
    picMakeProposal.Visible = False
    frmButton.Enabled = True
End Sub

Public Sub receiveProposal(strTrade As String)
    If PlayerNumber = Split(strTrade, ";")(2) Then
        tradingWith = Split(strTrade, ";")(1)
        outOfJailTradeA = Split(strTrade, ";")(3)
        outOfJailTradeB = Split(strTrade, ";")(4)
        txtTradeAmount2 = Split(strTrade, ";")(5)
        txtTradeAmount1 = Split(strTrade, ";")(6)
        For z = 1 To 20
            propertiesTrade1(z) = 0
            propertiesTrade2(z) = 0
        Next
        z = 1
        Do
            If Split(Split(strTrade, ";")(7), ",")(z) <> "TP1-END" Then
                propertiesTrade2(z) = Split(Split(strTrade, ";")(7), ",")(z)
                z = z + 1
            End If
        Loop Until Split(Split(strTrade, ";")(7), ",")(z) = "TP1-END"
        z = 1
        Do
            If Split(Split(strTrade, ";")(8), ",")(z) <> "TP2-END" Then
                propertiesTrade1(z) = Split(Split(strTrade, ";")(8), ",")(z)
                z = z + 1
            End If
        Loop Until Split(Split(strTrade, ";")(8), ",")(z) = "TP2-END"

        cmdProposeTrade.Visible = False
        cmdDoneTrade.Visible = False
        cmdCancelTrade.Visible = False
        cmdCounterTrade.Visible = True
        cmdRejectTrade.Visible = True
        cmdAcceptTrade.Visible = True
        picMakeProposal.Visible = True
        picMakeProposal.ZOrder
        frmButton.Enabled = False
        Call loadTradingProperties(Int(Split(strTrade, ";")(1)))
        lblTradeAnounce.Caption = "You received proposal from " & player(tradingWith).PName
    End If
End Sub

Private Sub cmdCounterTrade_Click()
    cmdProposeTrade.value = True
End Sub

Private Sub cmdAcceptTrade_Click()
    Dim strTrading As String
    strTrading = "cmd;acceptTrade;" & tableID & ";" & PlayerNumber
    strTrading = strTrading & ";" & tradingWith & ";" & outOfJailTradeA & ";" & outOfJailTradeB
    strTrading = strTrading & ";" & txtTradeAmount1 & ";" & txtTradeAmount2
    strTrading = strTrading & ";TP1"
    For z = 1 To 20
        If propertiesTrade1(z) <> 0 Then
            strTrading = strTrading & "," & propertiesTrade1(z)
        End If
    Next
    strTrading = strTrading & ",TP1-END;TPB"
    For z = 1 To 20
        If propertiesTrade2(z) <> 0 Then
            strTrading = strTrading & "," & propertiesTrade2(z)
        End If
    Next
    strTrading = strTrading & ",TP2-END;"
    frmMain.wskClient.SendData strTrading
    DoEvents
    picMakeProposal.Visible = False
    frmButton.Enabled = True
End Sub

Public Sub acceptTrade(strTrade As String)
    tradingWith = Split(strTrade, ";")(1)
    outOfJailTradeA = Split(strTrade, ";")(3)
    outOfJailTradeB = Split(strTrade, ";")(4)
    
    Call setPlayerStatus("cash", Int(Split(strTrade, ";")(1)), player(Split(strTrade, ";")(1)).cash - Int(Split(strTrade, ";")(5)))
    Call setPlayerStatus("cash", Int(Split(strTrade, ";")(2)), player(Split(strTrade, ";")(2)).cash + Int(Split(strTrade, ";")(5)))
    Call setPlayerStatus("cash", Int(Split(strTrade, ";")(2)), player(Split(strTrade, ";")(2)).cash - Int(Split(strTrade, ";")(6)))
    Call setPlayerStatus("cash", Int(Split(strTrade, ";")(1)), player(Split(strTrade, ";")(1)).cash + Int(Split(strTrade, ";")(6)))
    For z = 1 To 20
        propertiesTrade1(z) = 0
        propertiesTrade2(z) = 0
    Next
    z = 1
    Do
        If Split(Split(strTrade, ";")(7), ",")(z) <> "TP1-END" Then
            propertiesTrade2(z) = Split(Split(strTrade, ";")(7), ",")(z)
            slot(propertiesTrade2(z)).ownerPos = Split(strTrade, ";")(2)
            shpSlot(propertiesTrade2(z)).BackColor = picPlayer(slot(propertiesTrade2(z)).ownerPos).BackColor
            z = z + 1
        End If
    Loop Until Split(Split(strTrade, ";")(7), ",")(z) = "TP1-END"
    z = 1
    Do
        If Split(Split(strTrade, ";")(8), ",")(z) <> "TP2-END" Then
            propertiesTrade1(z) = Split(Split(strTrade, ";")(8), ",")(z)
            slot(propertiesTrade1(z)).ownerPos = Split(strTrade, ";")(1)
            shpSlot(propertiesTrade1(z)).BackColor = picPlayer(slot(propertiesTrade1(z)).ownerPos).BackColor
            z = z + 1
        End If
    Loop Until Split(Split(strTrade, ";")(8), ",")(z) = "TP2-END"

    If PlayerNumber = Split(strTrade, ";")(2) Then
        picMakeProposal.Visible = True
        picMakeProposal.ZOrder
        frmButton.Enabled = False
        Call loadTradingProperties(Int(Split(strTrade, ";")(1)))

        lblTradeAnounce.Caption = "Your proposal has been ACCEPTED by " & player(tradingWith).PName
        cmdProposeTrade.Visible = False
        cmdDoneTrade.Visible = True
        cmdCancelTrade.Visible = False
        cmdCounterTrade.Visible = False
        cmdRejectTrade.Visible = False
        cmdAcceptTrade.Visible = False
     End If
End Sub

Private Sub cmdRejectTrade_Click()
    frmMain.wskClient.SendData "cmd;rejectTrade;" & tableID & ";" & PlayerNumber & ";" & tradingWith & ";"
    DoEvents
    picMakeProposal.Visible = False
    frmButton.Enabled = True
End Sub
Public Sub rejectTrade(strTrade As String)
    tradingWith = Split(strTrade, ";")(1)
    If PlayerNumber = Split(strTrade, ";")(2) Then
        picMakeProposal.Visible = True
        picMakeProposal.ZOrder
        frmButton.Enabled = False
        Call loadTradingProperties(Int(Split(strTrade, ";")(1)))
        
        lblTradeAnounce.Caption = "Your proposal has been REJECTED by " & player(tradingWith).PName
        cmdProposeTrade.Visible = False
        cmdDoneTrade.Visible = True
        cmdCancelTrade.Visible = False
        cmdCounterTrade.Visible = False
        cmdRejectTrade.Visible = False
        cmdAcceptTrade.Visible = False
     End If
End Sub
Private Sub cmdDoneTrade_Click()
    picMakeProposal.Visible = False
    frmButton.Enabled = True
End Sub

Private Sub cmdCancelTrade_Click()
    picMakeProposal.Visible = False
    frmButton.Enabled = True
End Sub

Private Sub cmdAuction_Click()
    frmMain.wskClient.SendData "cmd;auction;" & tableID & ";" & PlayerNumber & ";" & player(PlayerNumber).currentSlot & ";"
    DoEvents
End Sub

Public Sub auction(number As Integer, cardNo As Integer)
    picAuction.Visible = True
    picAuction.ZOrder
    frmButton.Enabled = False
    cmdBuy.Enabled = False
    cmdAuction.Enabled = False
    lblComment.Caption = "Please start your bid!"
    tempBidCard = cardNo
    For z = cmdBidCash.LBound To cmdBidCash.UBound
        cmdBidCash(z).Enabled = True
    Next
    imgAuctionCard.Picture = LoadPicture("images/deed/deed" & cardNo & ".jpg")
    For z = 1 To maxPlayer
        If player(z).PID <> 0 Then
            picAuctionPlayer(z).Visible = True
            picAuctionPlayer(z).BackColor = picPlayer(z).BackColor
            lblAucName(z).Caption = player(z).PName
            lblCashBalance(z).Caption = player(z).cash
            lblCurrentBid.Caption = 0
        Else
            picAuctionPlayer(z).Visible = False
        End If
    Next
End Sub

Private Sub cmdBidCash_Click(Index As Integer)
    Dim bidAmount As Currency
    For z = cmdBidCash.LBound To cmdBidCash.UBound
        cmdBidCash(z).Enabled = False
    Next
    Select Case Index
        Case 0: bidAmount = 1
        Case 1: bidAmount = 5
        Case 2: bidAmount = 10
        Case 3: bidAmount = 20
        Case 4: bidAmount = 50
        Case 5: bidAmount = 100
        Case 6: bidAmount = 200
    End Select
    frmMain.wskClient.SendData "cmd;auctionBid;" & tableID & ";" & PlayerNumber & ";" & bidAmount & ";"
    DoEvents
End Sub

Public Sub addBidAmount(number As Integer, amount As Currency)
    tmrAuctionDelay.Enabled = True
    timeDelay = 0
    auctionTimeCount = 0
    tempBider = number
    lblCurrentBid.Caption = Int(lblCurrentBid.Caption) + Int(amount)
    lblCashBalance(tempBider).Caption = Int(lblCashBalance(tempBider).Caption) - Int(amount)
    lblComment.Caption = player(number).PName & " bid!"
    For z = cmdBidCash.LBound To cmdBidCash.UBound
        cmdBidCash(z).Enabled = True
    Next
End Sub

Private Sub tmrAuctionDelay_Timer()
    timeDelay = timeDelay + 1
    If timeDelay > gameRules(currentRules).auctionTimeDelay Then
        auctionTimeCount = auctionTimeCount + 1
        If auctionTimeCount = 1 Then
            lblComment.Caption = "Knowing one!"
        ElseIf auctionTimeCount = 3 Then
            lblComment.Caption = "Knowing twice!"
        ElseIf auctionTimeCount = 5 Then
            lblComment.Caption = "Sold!"
            For z = cmdBidCash.LBound To cmdBidCash.UBound
                cmdBidCash(z).Enabled = False
            Next
        ElseIf auctionTimeCount = 6 Then
            Call setPlayerStatus("cash", tempBider, player(tempBider).cash - Int(lblCurrentBid.Caption))
            slot(tempBidCard).hasOwner = True
            slot(tempBidCard).ownerPos = tempBider
            shpSlot(tempBidCard).BackStyle = 1
            shpSlot(tempBidCard).BackColor = picPlayer(tempBider).BackColor
            shpSlot(tempBidCard).Visible = True
            loadPlayerInfo (currentPlayer)
            lblComment.Caption = player(tempBider).PName & " are the new owner of this properties."
        ElseIf auctionTimeCount = 9 Then
            tmrAuctionDelay.Enabled = False
            picAuction.Visible = False
            frmButton.Enabled = True
            If currentPlayer = PlayerNumber Then
                cmdDone.Enabled = True
            End If
        End If
    End If
End Sub

Public Sub chance(number As Integer)
    imgCard.Picture = LoadPicture("images/chance/chance" & rndChance(currentChanceCard) & ".jpg")
    Select Case rndChance(currentChanceCard)
        Case 1
            player(number).cardOutOfJailA = True
        Case 2
            Dim numOfHouses As Integer
            Dim numOfHotels As Integer
            Dim TotalCash As Currency
            numOfHouses = 0
            numOfHotels = 0
            For z = 1 To 40
                If slot(z).ownerPos = number Then
                    If slot(z).numOfHouses = gameRules(currentRules).housesPerHotel Then
                        numOfHouses = numOfHouses + slot(z).numOfHouses - 1
                    Else
                        numOfHouses = numOfHouses + slot(z).numOfHouses
                    End If
                    numOfHotels = numOfHotels + slot(z).numOfHotels
                End If
            Next
            TotalCash = (numOfHouses * 25) + (numOfHotels * 100)
            Call setPlayerStatus("cash", number, player(number).cash - TotalCash)
        Case 3
            Dim totalPaid As Currency
            totalPaid = 0
            For z = 1 To maxPlayer
                If player(z).PID <> 0 Then
                    If z <> number Then
                        If Not player(z).inParking Then
                            Call setPlayerStatus("cash", number, player(z).cash + 50)
                            totalPaid = totalPaid + 50
                        End If
                    End If
                End If
            Next
            Call setPlayerStatus("cash", number, player(number).cash - totalPaid)
        Case 4
            Call setPlayerStatus("cash", number, player(number).cash - 15)
        Case 5
            Call setPlayerStatus("cash", number, player(number).cash + 50)
        Case 6
            Call setPlayerStatus("cash", number, player(number).cash + 150)
        Case 7
            If player(number).currentSlot > 12 Then
                Call setPlayerStatus("cash", number, player(number).cash + gameRules(currentRules).salary)
            End If
            Call moveTokenTo(number, 12)
        Case 8
            If player(number).currentSlot > 6 Then
                Call setPlayerStatus("cash", number, player(number).cash + gameRules(currentRules).salary)
            End If
            Call moveTokenTo(number, 6)
        Case 9
            Call moveTokenTo(number, 1)
            Call setPlayerStatus("cash", number, player(number).cash + gameRules(currentRules).salary)
        Case 10
            player(number).inJail = True
            player(number).numTurnInJail = 0
            imgJail(number).Visible = player(number).inJail
            s = play("oh_no.wav")
            Call moveTokenTo(number, 11)
        Case 11
            Call moveToken(number, -3)
        Case 12
            Call moveTokenTo(number, 10)
        Case 13
            Call moveTokenTo(number, 35)
        Case 14
            If player(number).currentSlot <= 13 Or player(number).currentSlot > 29 Then
                If slot(13).hasOwner Then
                    Call setPlayerStatus("cash", number, player(number).cash - ((numDice1 + numDice2) * 10))
                    Call setPlayerStatus("cash", slot(13).ownerPos, player(slot(13).ownerPos).cash + ((numDice1 + numDice2) * 10))
                Else
                    cmdBuy.Enabled = True
                    cmdAuction.Enabled = True
                End If
                Call moveTokenTo(number, 13)
            Else
                If slot(29).hasOwner Then
                    Call setPlayerStatus("cash", number, player(number).cash - ((numDice1 + numDice2) * 10))
                    Call setPlayerStatus("cash", slot(29).ownerPos, player(slot(29).ownerPos).cash + ((numDice1 + numDice2) * 10))
                Else
                    cmdBuy.Enabled = True
                    cmdAuction.Enabled = True
                End If
                Call moveTokenTo(number, 29)
            End If
        Case 15, 16
            If player(number).currentSlot < 6 And player(number).currentSlot > 36 Then
                Call moveTokenTo(number, 6)
            ElseIf player(number).currentSlot > 6 And player(number).currentSlot < 16 Then
                Call moveTokenTo(number, 16)
            ElseIf player(number).currentSlot > 16 And player(number).currentSlot < 26 Then
                Call moveTokenTo(number, 26)
            Else
                Call moveTokenTo(number, 36)
            End If
            If slot(player(number).currentSlot).hasOwner Then
                Dim NumOfRR As Integer
                Dim TotalRental As Currency
                NumOfRR = 0
                For z = 6 To 36 Step 10
                    If slot(z).ownerPos = slot(player(number).currentSlot).ownerPos Then
                        NumOfRR = NumOfRR + 1
                    End If
                Next
                TotalRental = deed(player(number).currentSlot).rentHouse(NumOfRR - 1) * 2
                Call setPlayerStatus("cash", number, player(number).cash - TotalRental)
                Call setPlayerStatus("cash", slot(player(number).currentSlot).ownerPos, player(slot(player(number).currentSlot).ownerPos).cash + TotalRental)
            Else
                cmdBuy.Enabled = True
                cmdAuction.Enabled = True
            End If
    End Select
    If player(number).cash < 0 And number = PlayerNumber Then
        cmdDone.Enabled = False
        cmdMortgage.Visible = True
        cmdMortgage.Enabled = True
        cmdBankrupt.Visible = True
        cmdBankrupt.Enabled = True
    End If
    currentChanceCard = currentChanceCard + 1
    If currentChanceCard > 16 Then
        currentChanceCard = 1
    End If
End Sub

Public Sub community(number As Integer)
    imgCard.Picture = LoadPicture("images/community/community" & rndCommunity(currentCommunityCard) & ".jpg")
    Select Case rndCommunity(currentCommunityCard)
        Case 1
            player(number).cardOutOfJailB = True
        Case 2
            Dim numOfHouses As Integer
            Dim numOfHotels As Integer
            Dim TotalCash As Currency
            numOfHouses = 0
            numOfHotels = 0
            For z = 1 To 40
                If slot(z).ownerPos = number Then
                    numOfHouses = numOfHouses + slot(z).numOfHouses
                    numOfHotels = numOfHotels + slot(z).numOfHotels
                End If
            Next
            TotalCash = (numOfHouses * 40) + (numOfHotels * 115)
            Call setPlayerStatus("cash", number, player(number).cash - TotalCash)
        Case 3
            Dim totalReceived As Currency
            totalReceived = 0
            For z = 1 To maxPlayer
                If player(z).PID <> 0 Then
                    If z <> number Then
                        If Not player(z).inParking Then
                            Call setPlayerStatus("cash", number, player(z).cash - 50)
                            totalReceived = totalReceived + 50
                        End If
                    End If
                End If
            Next
            Call setPlayerStatus("cash", number, player(number).cash + totalReceived)
            player(number).cash = player(number).cash + totalReceived
        Case 4
            player(number).inJail = True
            player(number).numTurnInJail = 0
            imgJail(number).Visible = player(number).inJail
            s = play("oh_no.wav")
            Call moveTokenTo(number, 11)
        Case 5
            Call moveTokenTo(number, 1)
            Call setPlayerStatus("cash", number, player(number).cash + gameRules(currentRules).salary)
        Case 6
            Call setPlayerStatus("cash", number, player(number).cash - 50)
        Case 7
            Call setPlayerStatus("cash", number, player(number).cash - 100)
        Case 8
            Call setPlayerStatus("cash", number, player(number).cash - 150)
        Case 9
            Call setPlayerStatus("cash", number, player(number).cash + 10)
        Case 10
            Call setPlayerStatus("cash", number, player(number).cash + 20)
        Case 11
            Call setPlayerStatus("cash", number, player(number).cash + 25)
        Case 12
            Call setPlayerStatus("cash", number, player(number).cash + 45)
        Case 13, 14, 15
            Call setPlayerStatus("cash", number, player(number).cash + 100)
        Case 16
            Call setPlayerStatus("cash", number, player(number).cash + 200)
    End Select
    If player(number).cash < 0 And number = PlayerNumber Then
        cmdDone.Enabled = False
        cmdMortgage.Visible = True
        cmdMortgage.Enabled = True
        cmdBankrupt.Visible = True
        cmdBankrupt.Enabled = True
    End If
    currentCommunityCard = currentCommunityCard + 1
    If currentCommunityCard > 16 Then
        currentCommunityCard = 1
    End If
End Sub

Private Sub cmdBuild_Click()
    Dim bool As Boolean
    bool = False
    If slot(viewSlot).numOfHouses < gameRules(currentrule).housesPerHotel Then
        If player(PlayerNumber).cash >= deed(viewSlot).houseCost Then
            bool = True
        End If
    Else
        If player(PlayerNumber).cash >= deed(viewSlot).hotelCost Then
            bool = True
        End If
    End If
    If housesAvailable + ((gameRules(currentRules).totalHotels - hotelsAvailable) * 5) > 0 Then
        If bool Then
            frmMain.wskClient.SendData "cmd;buildHouse;" & tableID & ";" & PlayerNumber & ";" & viewSlot & ";"
            DoEvents
            cmdBuild.Enabled = False
            cmdBuild.Visible = False
        Else
            MsgBox "You don't have enough money to build house", vbExclamation + vbOKOnly, "Build house"
        End If
    Else
        MsgBox "Sorry! The bank has no more house to build.", vbExclamation + vbOKOnly, "Build House"
    End If
End Sub

Public Sub buildHouse(number As Integer, slotNum As Integer)
    slot(slotNum).numOfHouses = slot(slotNum).numOfHouses + 1
    gameRules(currentRules).totalHouses = gameRules(currentRules).totalHouses - 1
    imgHouse1(slotNum).Visible = False
    imgHouse2(slotNum).Visible = False
    imgHouse3(slotNum).Visible = False
    imgHouse4(slotNum).Visible = False
    imgHotel(slotNum).Visible = False
    If slot(slotNum).numOfHouses < gameRules(currentRules).housesPerHotel Then
        Call setPlayerStatus("cash", number, player(number).cash - deed(slotNum).houseCost)
        If slot(slotNum).numOfHouses >= 1 Then imgHouse1(slotNum).Visible = True
        If slot(slotNum).numOfHouses >= 2 Then imgHouse2(slotNum).Visible = True
        If slot(slotNum).numOfHouses >= 3 Then imgHouse3(slotNum).Visible = True
        If slot(slotNum).numOfHouses >= 4 Then imgHouse4(slotNum).Visible = True
    Else
        gameRules(currentRules).totalHotels = gameRules(currentRules).totalHotels - 1
        slot(slotNum).numOfHotels = 1
        If slot(slotNum).numOfHouses = gameRules(currentRules).housesPerHotel Then imgHotel(slotNum).Visible = True
        Call setPlayerStatus("cash", number, player(number).cash - deed(slotNum).hotelCost)
    End If
End Sub

Private Sub cmdSell_Click()
    frmMain.wskClient.SendData "cmd;sellHouse;" & tableID & ";" & PlayerNumber & ";" & viewSlot & ";"
    DoEvents
    cmdSell.Enabled = False
    cmdSell.Visible = False
End Sub

Public Sub sellHouse(number As Integer, slotNum As Integer)
    slot(slotNum).numOfHouses = slot(slotNum).numOfHouses - 1
    imgHouse1(slotNum).Visible = False
    imgHouse2(slotNum).Visible = False
    imgHouse3(slotNum).Visible = False
    imgHouse4(slotNum).Visible = False
    imgHotel(slotNum).Visible = False
    If slot(slotNum).numOfHouses < gameRules(currentRules).housesPerHotel Then
        Call setPlayerStatus("cash", number, player(number).cash - deed(slotNum).houseCost)
        If slot(slotNum).numOfHouses >= 1 Then imgHouse1(slotNum).Visible = True
        If slot(slotNum).numOfHouses >= 2 Then imgHouse2(slotNum).Visible = True
        If slot(slotNum).numOfHouses >= 3 Then imgHouse3(slotNum).Visible = True
        If slot(slotNum).numOfHouses >= 4 Then imgHouse4(slotNum).Visible = True
    Else
        slot(slotNum).numOfHotels = 0
        Call setPlayerStatus("cash", number, player(number).cash - deed(slotNum).hotelCost)
    End If
End Sub

Private Sub cmdBuy_Click()
    frmMain.wskClient.SendData "cmd;buyProperties;" & tableID & ";" & PlayerNumber & ";" & player(PlayerNumber).currentSlot & ";EOT;"
    DoEvents
End Sub

Public Sub buyProperties(number As Integer, deedID As Integer)
    Call setPlayerStatus("cash", number, player(number).cash - deed(deedID).price)
    slot(deedID).hasOwner = True
    slot(deedID).ownerPos = number
    shpSlot(deedID).BackStyle = 1
    shpSlot(deedID).BackColor = picPlayer(number).BackColor
    shpSlot(deedID).Visible = True
    cmdBuy.Enabled = False
    cmdAuction.Enabled = False
    cmdDone.Enabled = False

    If number = PlayerNumber Then
        cmdDone.Enabled = True
    End If
    loadPlayerInfo (currentPlayer)
End Sub

Private Sub cmdMortgage_Click()
    frmMain.wskClient.SendData "cmd;mortgage;" & tableID & ";" & PlayerNumber
    DoEvents
End Sub

Public Sub mortgage(number As Integer)
    Dim tmpLeft As Long
    Dim tmpTop As Long
    Dim deedCount As Integer
    picMortgage.Visible = True
    picMortgage.ZOrder
    frmButton.Enabled = False
    deedCount = 0
    tmpTop = 400
    tmpLeft = 100
    cmdMortgageDC.Enabled = False
    cmdUnmortgageDC.Enabled = False
    If PlayerNumber = number Then
        cmdMenu.Enabled = False
        cmdClose.Enabled = True
    Else
        cmdMenu.Enabled = True
        cmdClose.Enabled = False
    End If
    For i = 1 To 40
        If deed(i).mortgageValue > 0 Then
            If slot(i).hasOwner And slot(i).ownerPos = number Then
                If Not slot(i).onMortgage Then
                    imgMortgageDeedCard(i).Picture = LoadPicture("images/deed/deed" & i & ".jpg")
                Else
                    imgMortgageDeedCard(i).Picture = LoadPicture("images/deed/deedM" & i & ".jpg")
                End If
                imgMortgageDeedCard(i).Top = tmpTop
                imgMortgageDeedCard(i).Left = tmpLeft
                imgMortgageDeedCard(i).Visible = True
                imgMortgageDeedCard(i).ZOrder
            Else
                imgMortgageDeedCard(i).Visible = False
            End If
            tmpTop = tmpTop + 550
            deedCount = deedCount + 1
        End If
        If deedCount > 10 Then
            tmpTop = 400
            tmpLeft = tmpLeft + 2500
            deedCount = 0
        End If
    Next
End Sub

Private Sub imgMortgageDeedCard_Click(Index As Integer)
    If PlayerNumber = currentPlayer Then
        cmdMortgageDC.Visible = True
        cmdUnmortgageDC.Visible = True
        cmdMortgageDC.Enabled = False
        cmdUnmortgageDC.Enabled = False
        
        If selectedCard > 0 Then
            imgMortgageDeedCard(selectedCard).Appearance = 0
        End If
        selectedCard = Index
        imgMortgageDeedCard(selectedCard).Appearance = 1
        If Not slot(selectedCard).onMortgage Then
            imgCardPreview.Picture = LoadPicture("images/deed/deed" & selectedCard & ".jpg")
        Else
            imgCardPreview.Picture = LoadPicture("images/deed/deedM" & selectedCard & ".jpg")
        End If
        imgCardPreview.Visible = True
        If slot(selectedCard).onMortgage Then
            cmdUnmortgageDC.Enabled = True
        ElseIf slot(selectedCard).numOfHouses = 0 Then
            cmdMortgageDC.Enabled = True
        End If
    End If
End Sub

Private Sub cmdMortgageDC_Click()
    frmMain.wskClient.SendData "cmd;mortgageDeedCard;" & tableID & ";" & PlayerNumber & ";" & selectedCard
    DoEvents
End Sub

Public Sub mortgageDeedCard(number As Integer, cardNo As Integer)
    Call setPlayerStatus("cash", number, player(number).cash + deed(cardNo).mortgageValue)
    slot(cardNo).onMortgage = True
    imgMortgageDeedCard(cardNo).Picture = LoadPicture("images/deed/deedM" & cardNo & ".jpg")
    If PlayerNumber = number Then
        imgMortgageDeedCard(cardNo).Appearance = 0
        cmdMortgageDC.Enabled = False
        selectedCard = 0
    End If
End Sub

Private Sub cmdUnmortgageDC_Click()
    frmMain.wskClient.SendData "cmd;unmortgageDeedCard;" & tableID & ";" & PlayerNumber & ";" & selectedCard
    DoEvents
End Sub

Public Sub unmortgageDeedCard(number As Integer, cardNo As Integer)
    Call setPlayerStatus("cash", number, player(number).cash - deed(cardNo).mortgageValue - (deed(cardNo).mortgageValue * 0.1))
    slot(cardNo).onMortgage = False
    imgMortgageDeedCard(cardNo).Picture = LoadPicture("images/deed/deed" & cardNo & ".jpg")
    If PlayerNumber = number Then
        imgMortgageDeedCard(cardNo).Appearance = 0
        cmdMortgageDC.Enabled = False
        selectedCard = 0
    End If
End Sub

Private Sub cmdClose_Click()
    frmMain.wskClient.SendData "cmd;mortgageClose;" & tableID & ";" & PlayerNumber & ";"
    DoEvents
End Sub

Public Sub closeMortgage()
    cmdMenu.Enabled = True
    picMortgage.Visible = False
    If player(currentPlayer).cash > 0 Then
        cmdBankrupt.Visible = False
        cmdDone.Enabled = True
    End If
    frmButton.Enabled = True
End Sub

Public Sub playerquit(number As Integer)
    Dim count As Integer
    If gameStarted Then
        MsgBox player(number).PName & " has quit the game.", vbInformation + vbOKOnly, "Player Quit"
        imgPlayerToken(number).Picture = LoadPicture("")
        imgJail(number).Visible = False
        lblPlayerCash(number).Caption = ""
    End If
    Call setPlayerStatus("PID", number, 0)
    Call setPlayerStatus("PName", number, "")
    
    imgWPToken(number).Picture = LoadPicture("")
    lblWPJoinName(number).Caption = ""
    lblWReady(number).Caption = ""
    count = 0
    For z = 1 To maxPlayer
        If player(z).PID <> 0 Then
            count = count + 1
        End If
    Next
    If Not gameStarted And player(number).keyPlayer Then
        Unload Me
    End If
    If count <= 1 And gameStarted Then
        MsgBox "Congratulation! You are the winner", vbInformation + vbOKOnly, "Game"
    End If
End Sub
'############### In game function  - End

Private Sub imgSlot_Click(Index As Integer)
    Dim colorGroupCount As Integer
    Dim ownerGroupCount As Integer
    Dim housesCount1 As Integer
    Dim housesCount2 As Integer
    Dim canBuild As Boolean
    Dim canSell As Boolean
    On Error GoTo unloadImage
    viewSlot = Index
    imgDeedCard.Picture = LoadPicture("images/deed/deed" & Index & ".jpg")
    cmdBuild.Visible = False
    cmdBuild.Enabled = False
    cmdSell.Visible = False
    cmdSell.Enabled = False
    If gameStarted And slot(Index).hasOwner And slot(Index).ownerPos = PlayerNumber And deed(Index).color <> "none" Then
        colorGroupCount = 0
        ownerGroupCount = 0
        canBuild = True
        canSell = True
        housesCount1 = slot(Index).numOfHouses
        For z = 1 To 40
            If deed(z).mortgageValue > 0 And deed(z).color <> "none" Then
                If deed(z).color = deed(Index).color Then
                    colorGroupCount = colorGroupCount + 1
                End If
                If slot(z).hasOwner And slot(z).ownerPos = slot(Index).ownerPos Then
                    If slot(Index).numOfHouses - slot(z).numOfHouses >= 1 Or slot(Index).onMortgage Then
                        canBuild = False
                    End If
                    If slot(Index).numOfHouses > 0 And (slot(Index).numOfHouses - slot(z).numOfHouses < 0 Or slot(Index).numOfHouses - slot(z).numOfHouses > 1) Then
                        canSell = False
                    End If
                    ownerGroupCount = ownerGroupCount + 1
                End If
            End If
        Next
        If ownerGroupCount = colorGroupCount Then
            If canBuild Then
                cmdBuild.Visible = True
                cmdBuild.Enabled = True
            End If
            If canSell Then
                cmdSell.Visible = True
                cmdSell.Enabled = True
            End If
        End If
    End If
    Exit Sub
unloadImage:
    imgDeedCard.Picture = LoadPicture("")
End Sub

Public Sub loadPlayerInfo(number As Integer)
    If player(number).cardOutOfJailA Then
        imgJailA.Picture = LoadPicture("images/chance/chance1.jpg")
        imgJailA.Visible = True
    Else
        imgJailA.Picture = LoadPicture("")
        imgJailA.Visible = False
    End If
    If player(number).cardOutOfJailB Then
        imgJailB.Picture = LoadPicture("images/community/community1.jpg")
        imgJailB.Visible = True
    Else
        imgJailB.Picture = LoadPicture("")
        imgJailB.Visible = False
    End If
End Sub


Private Sub tmrMusic_Timer()
    On Error Resume Next
    If playMusicBool Then
        If v_dmss.GetSeek >= v_dms.GetLength Then
            CloseMidi
            PlayMidi (0)
        End If
    End If
End Sub

Private Sub txtSendTableMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim$(txtSendTableMsg) <> "" Then
            frmMain.wskClient.SendData "msg;table;" & tableID & ";" & PlayerNumber & ";" & txtSendTableMsg
            DoEvents
            txtSendTableMsg.Text = ""
            txtSendTableMsg.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Public Sub addChatMsg(strmsg As String)
    txtTblChatMsg.Text = txtTblChatMsg.Text & vbCrLf & Right(strmsg, Len(strmsg) - (Len(Split(strmsg, ";")(0)) + 1))
    txtTblChatMsg.SelStart = Len(txtTblChatMsg.Text)
End Sub

Private Sub txtTradeAmount1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 32               ' Control keys are OK.
        Case 48 To 57              ' This is a digit.
        Case Else                  ' Reject any other key.
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTradeAmount1_Validate(Cancel As Boolean)
    If Int(txtTradeAmount1.Text) > player(PlayerNumber).cash Then
        MsgBox "The player do not have that much money", vbInformation + vbOKCancel, "Trade"
        txtTradeAmount1.Text = 0
        txtTradeAmount1.SetFocus
    End If
End Sub

Private Sub txtTradeAmount2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 32               ' Control keys are OK.
        Case 48 To 57              ' This is a digit.
        Case Else                  ' Reject any other key.
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTradeAmount2_Validate(Cancel As Boolean)
    If Int(txtTradeAmount1.Text) > player(tradingWith).cash Then
        MsgBox "The player do not have that much money", vbInformation + vbOKCancel, "Trade"
        txtTradeAmount2.Text = 0
        txtTradeAmount2.SetFocus
    End If
End Sub

Private Sub Form_Resize()
    picWaiting.Left = Me.Width / 2 - picWaiting.Width / 2
    picWaiting.Top = Me.Height / 2 - picWaiting.Height / 2
    picMortgage.Left = Me.Width / 2 - picMortgage.Width / 2
    picMortgage.Top = Me.Height / 2 - picMortgage.Height / 2
    picAuction.Left = Me.Width / 2 - picAuction.Width / 2
    picAuction.Top = Me.Height / 2 - picAuction.Height / 2
    picSelectTrader.Left = Me.Width / 2 - picSelectTrader.Width / 2
    picSelectTrader.Top = Me.Height / 2 - picSelectTrader.Height / 2
    picMakeProposal.Left = Me.Width / 2 - picMakeProposal.Width / 2
    picMakeProposal.Top = Me.Height / 2 - picMakeProposal.Height / 2
    picStatus.Left = Me.Width / 2 - picStatus.Width / 2
    picStatus.Top = Me.Height / 2 - picStatus.Height / 2
End Sub

Private Sub Form_Activate()
    activeWindow.Show
    activeWindow.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.tmrMusic.Enabled = True
    tmrMusic.Enabled = False
    Call ResetPlayerStatus
    frmMain.Show
    Set activeWindow = preActiveWindow
    frmMain.wskClient.SendData "status;quitGame;" & tableID & ";" & PlayerNumber & ";"
    DoEvents
End Sub

