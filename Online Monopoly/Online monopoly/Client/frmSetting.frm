VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Monopoly - Setting"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetting.frx":0000
   ScaleHeight     =   3045
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Picture         =   "frmSetting.frx":190C8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Multimedia"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin MSComctlLib.Slider Volume 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.CheckBox chkSound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Play Sound"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkMusic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Play Music"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Music Volume"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkMusic_Click()
    If chkMusic.value = 1 Then
        If Not playMusicBool Then
            PlayMidi (1)
        End If
        playMusicBool = True
    Else
        If playMusicBool Then CloseMidi
        playMusicBool = False
    End If
    SaveSetting "Monopoly", "Game Setting", "Music", playMusicBool
End Sub

Private Sub chkSound_Click()
    If chkSound.value = 1 Then
        playSoundBool = True
    Else
        playSoundBool = False
    End If
    SaveSetting "Monopoly", "Game Setting", "Sound", playSoundBool
End Sub

Private Sub cmdClose_Click()
    Set activeWindow = preActiveWindow
    Unload Me
End Sub

Private Sub Form_Load()
    Set preActiveWindow = activeWindow
    Set activeWindow = Me
    If playMusicBool Then
        chkMusic.value = 1
    Else
        chkMusic.value = 0
    End If
    If playSoundBool Then
        chkSound.value = 1
    Else
        chkSound.value = 0
    End If
    
    Volume.value = (GetSetting("Monopoly", "Game Setting", "Music Volume", 0) + 1500) / 25
End Sub

Private Sub Volume_Scroll()
    Call v_dmp.SetMasterVolume((Volume.value * 25) - 1500)
    SaveSetting "Monopoly", "Game Setting", "Music Volume", v_dmp.GetMasterVolume
End Sub
