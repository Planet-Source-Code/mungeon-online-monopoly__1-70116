VERSION 5.00
Begin VB.Form frmTrading 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monopoly"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlayerName 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Height          =   375
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdPlayerName 
      BackColor       =   &H0080FF80&
      Height          =   375
      Index           =   3
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdPlayerName 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Index           =   4
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdPlayerName 
      BackColor       =   &H00FF8080&
      Height          =   375
      Index           =   2
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancelTrading 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frmTrading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
