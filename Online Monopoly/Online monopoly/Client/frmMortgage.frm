VERSION 5.00
Begin VB.Form frmMortgage 
   Caption         =   "Monopoly"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
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
   ScaleHeight     =   7110
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnmortgage 
      Caption         =   "unmortgage"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdMortgage 
      Caption         =   "mortgage"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "done"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label lblCash 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "your cash : "
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgDeedCard 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmMortgage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selectedCard As Integer
Dim cash As Currency
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdMortgage_Click()
    cash = cash + deed(selectedCard).mortgageValue
    lblCash.Caption = cash
    slot(selectedCard).onMortgage = True
    imgDeedCard(selectedCard).Picture = LoadPicture("images/deed/deedM19.jpg")
    imgDeedCard(selectedCard).Appearance = 0
    cmdMortgage.Enabled = False
    selectedCard = 0
End Sub

Private Sub cmdUnMortgage_Click()

End Sub

Private Sub Form_Load()
    Dim imgCount As Integer
    Set activeWindow = Me
    cash = player(PlayerNumber).cash
    lblCash.Caption = cash
    cmdMortgage.Enabled = False
    cmdUnmortgage.Enabled = False
    imgCount = 1
    For i = 1 To 40
        Load imgDeedCard(i)
        imgDeedCard(i).Left = imgDeedCard(i - 1).Left + 400
        If slot(i).hasOwner And slot(i).ownerPos = PlayerNumber Then
            If Not slot(i).onMortgage Then
                imgDeedCard(i).Picture = LoadPicture("images/deed/deed" & i & ".jpg")
            Else
                'imgdeedCard(imgCount).Picture = LoadPicture("images/deed/deed" & i & ".jpg")
            End If
            imgDeedCard(i).Visible = True
            imgDeedCard(i).ZOrder
        End If
        If i > 16 Then
            imgDeedCard(i).Left = 400
            imgDeedCard(i).Top = imgDeedCard(i - 1).Top + 2000
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set activeWindow = Monopoly
End Sub

Private Sub imgDeedCard_Click(Index As Integer)
    selectedCard = Index
    cmdMortgage.Visible = True
    cmdUnmortgage.Visible = True
    cmdMortgage.Enabled = False
    cmdUnmortgage.Enabled = False
    For i = 1 To 40
        If i = selectedCard Then
            imgDeedCard(i).Appearance = 1
            If slot(i).onMortgage Then
                cmdUnmortgage.Enabled = True
            ElseIf slot(i).numOfHotels = 0 And slot(i).numOfHouses = 0 Then
                cmdMortgage.Enabled = True
            End If
        Else
            imgDeedCard(i).Appearance = 0
        End If
    Next
End Sub
