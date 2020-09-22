VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Fortress - Main"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form3"
   ScaleHeight     =   6150
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   465
      Left            =   3150
      TabIndex        =   2
      Top             =   5550
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   465
      Left            =   1800
      TabIndex        =   1
      Top             =   5550
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   465
      Left            =   300
      TabIndex        =   0
      Top             =   5550
      Width           =   915
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save3:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1350
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
      Visible         =   0   'False
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save2:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1350
      TabIndex        =   5
      Top             =   1650
      Width           =   1815
      Visible         =   0   'False
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save1:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1350
      TabIndex        =   4
      Top             =   900
      Width           =   1815
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4365
      Visible         =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "Form3.frx":0000
      Top             =   0
      Width           =   4290
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form1
Form1.Show
Form1.Timer3.Enabled = True
Unload Form3
End Sub

Private Sub Command2_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Open (App.Path & "\Data\Save1.for") For Input As #1
    Line Input #1, OCannonHealth
    Line Input #1, OCannonmax
    Line Input #1, OCannonLevel
    Line Input #1, OHouseHealth
    Line Input #1, OHouseMax
    Line Input #1, OHouseLevel
    Line Input #1, OGold
    Line Input #1, OTotalScore
    Line Input #1, OLevels
    Label2.Caption = "Save1: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
    If OLevels = 0 Then
    Label2.Enabled = False
    End If
Close #1
Open (App.Path & "\Data\Save2.for") For Input As #1
    Line Input #1, OCannonHealth
    Line Input #1, OCannonmax
    Line Input #1, OCannonLevel
    Line Input #1, OHouseHealth
    Line Input #1, OHouseMax
    Line Input #1, OHouseLevel
    Line Input #1, OGold
    Line Input #1, OTotalScore
    Line Input #1, OLevels
    Label3.Caption = "Save3: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
    If OLevels = 0 Then
    Label3.Enabled = False
    End If
Close #1
Open (App.Path & "\Data\Save3.for") For Input As #1
    Line Input #1, OCannonHealth
    Line Input #1, OCannonmax
    Line Input #1, OCannonLevel
    Line Input #1, OHouseHealth
    Line Input #1, OHouseMax
    Line Input #1, OHouseLevel
    Line Input #1, OGold
    Line Input #1, OTotalScore
    Line Input #1, OLevels
    Label4.Caption = "Save3: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
    If OLevels = 0 Then
    Label4.Enabled = False
    End If
Close #1
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Label2_Click()
Load Form1
Form1.Hide
Open (App.Path & "\Data\Save1.for") For Input As #1
    Line Input #1, OCannonHealth
    CannonHealth = OCannonHealth
    Line Input #1, OCannonmax
    CannonMax = OCannonmax
    Line Input #1, OCannonLevel
    CannonLevel = OCannonLevel
    Line Input #1, OHouseHealth
    HouseHealth = OHouseHealth
    Line Input #1, OHouseMax
    HouseMax = OHouseMax
    Line Input #1, OHouseLevel
    HouseLevel = OHouseLevel
    Line Input #1, OGold
    Gold = OGold
    Line Input #1, OTotalScore
    TotalScore = OTotalScore
    Line Input #1, OLevels
    Levels = OLevels
Close #1
Form2.Show
Unload Form3
End Sub

Private Sub Label3_Click()
Load Form1
Form1.Hide
Open (App.Path & "\Data\Save2.for") For Input As #1
    Line Input #1, OCannonHealth
    CannonHealth = OCannonHealth
    Line Input #1, OCannonmax
    CannonMax = OCannonmax
    Line Input #1, OCannonLevel
    CannonLevel = OCannonLevel
    Line Input #1, OHouseHealth
    HouseHealth = OHouseHealth
    Line Input #1, OHouseMax
    HouseMax = OHouseMax
    Line Input #1, OHouseLevel
    HouseLevel = OHouseLevel
    Line Input #1, OGold
    Gold = OGold
    Line Input #1, OTotalScore
    TotalScore = OTotalScore
    Line Input #1, OLevels
    Levels = OLevels
Close #1
Form2.Show
Unload Form3
End Sub

Private Sub Label4_Click()
Load Form1
Form1.Hide
Open (App.Path & "\Data\Save3.for") For Input As #1
    Line Input #1, OCannonHealth
    CannonHealth = OCannonHealth
    Line Input #1, OCannonmax
    CannonMax = OCannonmax
    Line Input #1, OCannonLevel
    CannonLevel = OCannonLevel
    Line Input #1, OHouseHealth
    HouseHealth = OHouseHealth
    Line Input #1, OHouseMax
    HouseMax = OHouseMax
    Line Input #1, OHouseLevel
    HouseLevel = OHouseLevel
    Line Input #1, OGold
    Gold = OGold
    Line Input #1, OTotalScore
    TotalScore = OTotalScore
    Line Input #1, OLevels
    Levels = OLevels
Close #1
Form2.Show
Unload Form3
End Sub
