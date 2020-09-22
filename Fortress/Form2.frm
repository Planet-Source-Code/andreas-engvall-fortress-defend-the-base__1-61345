VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Fortress - Shop"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   ScaleHeight     =   3615
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Upgrade"
      Height          =   315
      Left            =   3450
      TabIndex        =   8
      Top             =   3150
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Repair"
      Height          =   315
      Left            =   3450
      TabIndex        =   7
      Top             =   2400
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Upgrade"
      Height          =   315
      Left            =   3450
      TabIndex        =   4
      Top             =   1350
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Repair"
      Height          =   315
      Left            =   3450
      TabIndex        =   2
      Top             =   600
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick  a slot to save to:"
      Height          =   315
      Left            =   2100
      TabIndex        =   14
      Top             =   300
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save3:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2100
      TabIndex        =   13
      Top             =   2400
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save2:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2100
      TabIndex        =   12
      Top             =   1650
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save1:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2100
      TabIndex        =   11
      Top             =   900
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   1950
      TabIndex        =   10
      Top             =   0
      Width           =   2715
      Visible         =   0   'False
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   1950
      Picture         =   "Form2.frx":0000
      Top             =   3300
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        x0"
      Height          =   345
      Index           =   2
      Left            =   1950
      TabIndex        =   9
      Top             =   3300
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cost for upgrade: "
      Height          =   465
      Index           =   1
      Left            =   3450
      TabIndex        =   6
      Top             =   2700
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Health:"
      Height          =   465
      Index           =   1
      Left            =   3450
      TabIndex        =   5
      Top             =   1950
      Width           =   1065
      Visible         =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cost for upgrade: "
      Height          =   465
      Index           =   0
      Left            =   3450
      TabIndex        =   3
      Top             =   900
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Health:"
      Height          =   465
      Index           =   0
      Left            =   3450
      TabIndex        =   1
      Top             =   150
      Width           =   1065
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   0
      Left            =   2100
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   1
      Left            =   2250
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   2
      Left            =   2400
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   3
      Left            =   2550
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   4
      Left            =   2700
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   0
      Left            =   2100
      Top             =   1800
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   1
      Left            =   2250
      Top             =   1800
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   2
      Left            =   2400
      Top             =   1800
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   3
      Left            =   2550
      Top             =   1800
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   4
      Left            =   2700
      Top             =   1800
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Image6 
      Height          =   1290
      Left            =   2100
      Picture         =   "Form2.frx":03D5
      Top             =   1950
      Width           =   1260
      Visible         =   0   'False
   End
   Begin VB.Image Image5 
      Height          =   1560
      Left            =   1950
      Picture         =   "Form2.frx":0E46
      Top             =   150
      Width           =   1470
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   1950
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Visible         =   0   'False
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   0
      Picture         =   "Form2.frx":1E60
      Top             =   2100
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   0
      Picture         =   "Form2.frx":25B1
      Top             =   1500
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   0
      Picture         =   "Form2.frx":2D4F
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "Form2.frx":351D
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If HouseHealth < HouseMax Then
    If Gold >= HouseMax - HouseHealth Then
        Gold = Gold - (HouseMax - HouseHealth)
        Label2(2).Caption = "       x" & Gold
        Dim Svar As Integer
        Svar = MsgBox("You've repaired the House with a cost of: " & HouseMax - HouseHealth & " Gold!", vbOKOnly, "You've repaired the house")
        HouseHealth = HouseMax
        Label2(0).Caption = "Health: " & HouseHealth & "/" & HouseMax
    Else
        Dim Svar1 As Integer
        Svar1 = MsgBox("You're " & Gold - (HouseMax - HouseHealth) & " short", vbOKOnly, "No money!")
    End If
Else
    Dim Svar2 As Integer
    Svar2 = MsgBox("You don't have to repair", vbOKOnly, "No need for repairing.")
End If
End Sub

Private Sub Command2_Click()
If HouseLevel < 5 Then
    If Gold >= HouseLevel * 2000 Then
        Gold = Gold - HouseLevel * 2000
        Label2(2).Caption = "       x" & Gold
        HouseLevel = HouseLevel + 1
        HouseMax = HouseLevel * 550
        HouseHealth = HouseHealth + 550
        Label2(0).Caption = "Health:   " & HouseHealth & " / " & HouseMax¨
        Label3(0).Caption = "Cost for upgrade: " & HouseLevel * 2000
        If HouseLevel = 2 Then Shape4(1).FillColor = vbRed
        If HouseLevel = 3 Then Shape4(2).FillColor = vbRed
        If HouseLevel = 4 Then Shape4(3).FillColor = vbRed
        If HouseLevel = 5 Then Shape4(4).FillColor = vbRed
    Else
        Dim Svar1 As Integer
        Svar1 = MsgBox("You're " & Gold - (HouseLevel * 2000) & " short", vbOKOnly, "No money!")
        If HouseLevel = 2 Then Shape4(1).BackColor = vbRed
    End If
Else
    Dim Svar As Integer
    Svar = MsgBox("You've got full level already, no need to upgrade!", vbOKOnly, "You can't upgrade!")
End If
End Sub

Private Sub Command3_Click()
If CannonHealth < CannonMax Then
    If Gold >= CannonMax - CannonHealth Then
        Gold = Gold - (CannonMax - CannonHealth)
        Label2(2).Caption = "       x" & Gold
        Dim Svar As Integer
        Svar = MsgBox("You've repaired the cannon with a cost of: " & CannonMax - CannonHealth & " Gold!", vbOKOnly, "You've repaired the cannon")
        CannonHealth = CannonMax
        Label2(1).Caption = "Health: " & CannonHealth & "/" & CannonMax
    Else
        Dim Svar1 As Integer
        Svar1 = MsgBox("You're " & Gold - (CannonMax - CannonHealth) & " short", vbOKOnly, "No money!")
    End If
Else
    Dim Svar2 As Integer
    Svar2 = MsgBox("You don't have to repair", vbOKOnly, "No need for repairing.")
End If
End Sub

Private Sub Command4_Click()
If CannonLevel < 5 Then
    If Gold >= CannonLevel * 3000 Then
        Gold = Gold - CannonLevel * 3000
        Label2(2).Caption = "       x" & Gold
        CannonLevel = CannonLevel + 1
        CannonMax = CannonLevel * 500
        CannonHealth = CannonHealth + 500
        Label2(1).Caption = "Health:   " & CannonHealth & "/" & CannonMax
        Label3(1).Caption = "Cost for upgrade: " & CannonLevel * 3000
        If CannonLevel = 2 Then Shape6(1).FillColor = vbRed
        If CannonLevel = 3 Then Shape6(2).FillColor = vbRed
        If CannonLevel = 4 Then Shape6(3).FillColor = vbRed
        If CannonLevel = 5 Then Shape6(4).FillColor = vbRed
    Else
        Dim Svar1 As Integer
        Svar1 = MsgBox("You're " & Gold - (CannonLevel * 3000) & " short", vbOKOnly, "No money!")
    End If
Else
    Dim Svar As Integer
    Svar = MsgBox("You've got full level already, no need to upgrade!", vbOKOnly, "You can't upgrade!")
End If
End Sub

Private Sub Image1_Click()
Save_false
Label2(0).Caption = "Health: " & HouseHealth & " / " & HouseMax
Label2(1).Caption = "Health: " & CannonHealth & " / " & CannonMax
For Number = 0 To Shape4.UBound
    Shape4(Number).Visible = True
Next
For Number1 = 0 To Shape6.UBound
    Shape6(Number1).Visible = True
Next
For Number2 = 0 To Label2.UBound
    Label2(Number2).Visible = True
Next
For Number3 = 0 To Label3.UBound
    Label3(Number3).Visible = True
Next
Label2(2).Caption = "       x" & Gold

Label3(0).Caption = "Cost for upgrade: " & HouseLevel * 2000
Label3(1).Caption = "Cost for upgrade: " & CannonLevel * 3000
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Label1.Visible = True
        If HouseLevel = 2 Then
            Shape4(1).FillColor = vbRed
        End If
        If HouseLevel = 3 Then
            Shape4(1).FillColor = vbRed
            Shape4(2).FillColor = vbRed
        End If
        If HouseLevel = 4 Then
            Shape4(1).FillColor = vbRed
            Shape4(2).FillColor = vbRed
            Shape4(3).FillColor = vbRed
        End If
        If HouseLevel = 5 Then
            Shape4(4).FillColor = vbRed
            Shape4(1).FillColor = vbRed
            Shape4(2).FillColor = vbRed
            Shape4(3).FillColor = vbRed
        End If

        If CannonLevel = 2 Then
            Shape6(1).FillColor = vbRed
        End If
        If CannonLevel = 3 Then
            Shape6(1).FillColor = vbRed
            Shape6(2).FillColor = vbRed
        End If
        If CannonLevel = 4 Then
            Shape6(1).FillColor = vbRed
            Shape6(2).FillColor = vbRed
            Shape6(3).FillColor = vbRed
        End If
        If CannonLevel = 5 Then
            Shape6(4).FillColor = vbRed
            Shape6(1).FillColor = vbRed
            Shape6(2).FillColor = vbRed
            Shape6(3).FillColor = vbRed
        End If
End Sub

Private Sub Image2_Click()
Shop_false
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True

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
    Label5.Caption = "Save1: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
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
    Label6.Caption = "Save2: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
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
    Label7.Caption = "Save3: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
Close #1
End Sub

Private Sub Image3_Click()
If CannonHealth >= 1 Then
    Form1.Image8.Visible = True
    Form1.Shape7.Visible = True
    Damage = CannonLevel * 20
    For Number = 0 To Form1.Shape6.UBound
      Form1.Shape6(Number).Visible = True
    Next
    Else
        Damage = 15
End If
Form1.Timer1.Enabled = False
Levels = Levels + 1
Form1.Enabled = True
Label1.Visible = True
If Levels < 6 Then
Form1.Image1.Picture = LoadPicture(App.Path & "\Pics\BG1Start.gif")
Else
Form1.Image1.Picture = LoadPicture(App.Path & "\Pics\BG2Start.gif")
End If
Form1.Image2.Visible = False
Form1.Image3.Visible = False
Form1.Image10.Visible = False
Form1.Image11.Visible = False
Form1.Image12.Visible = False
Form1.Image13.Visible = False
Form1.Image14.Visible = False
Form1.Shape9.Visible = False
Form1.Shape10.Visible = False
Form1.Shape11.Visible = False
Form1.Shape12.Visible = False
Form1.Shape13.Visible = False
Form1.Label5.Caption = "Level: " & Levels
Form1.Image4.Visible = False
Form1.Timer3.Enabled = True
Form1.Shape8.Width = HouseHealth * 1
Form1.Shape8.Visible = False
Form1.Shape7.Visible = False
Form1.Timer2.Enabled = True
Form1.Shape5.Visible = False
Form1.Shape1.Visible = False
Form1.Shape2.Visible = False
Form1.Shape3.Visible = False
Form1.Image8.Visible = False
Form1.Image5.Visible = False
For Number1 = 0 To Form1.Shape4.UBound
    Form1.Shape4(Number1).Visible = False
Next
For Number2 = 0 To Form1.Shape6.UBound
    Form1.Shape6(Number2).Visible = False
Next
TimeLeft = 40
If CannonHealth > 0 Then
    Form1.Shape7.Width = CannonHealth * 1
End If
Form1.Label2.Caption = "       x" & Gold
    For Number = 0 To Form1.Label4.UBound
        Form1.Label4(Number).Visible = False
    Next
Form1.Image9.Visible = False
        If HouseLevel = 2 Then
            Form1.Shape4(1).FillColor = vbRed
        End If
        If HouseLevel = 3 Then
            Form1.Shape4(1).FillColor = vbRed
            Form1.Shape4(2).FillColor = vbRed
        End If
        If HouseLevel = 4 Then
            Form1.Shape4(1).FillColor = vbRed
            Form1.Shape4(2).FillColor = vbRed
            Form1.Shape4(3).FillColor = vbRed
        End If
        If HouseLevel = 5 Then
            Form1.Shape4(4).FillColor = vbRed
            Form1.Shape4(1).FillColor = vbRed
            Form1.Shape4(2).FillColor = vbRed
            Form1.Shape4(3).FillColor = vbRed
        End If

        If CannonLevel = 2 Then
            Form1.Shape6(1).FillColor = vbRed
        End If
        If CannonLevel = 3 Then
            Form1.Shape6(1).FillColor = vbRed
            Form1.Shape6(2).FillColor = vbRed
        End If
        If CannonLevel = 4 Then
            Form1.Shape6(1).FillColor = vbRed
            Form1.Shape6(2).FillColor = vbRed
            Form1.Shape6(3).FillColor = vbRed
        End If
        If CannonLevel = 5 Then
            Form1.Shape6(4).FillColor = vbRed
            Form1.Shape6(1).FillColor = vbRed
            Form1.Shape6(2).FillColor = vbRed
            Form1.Shape6(3).FillColor = vbRed
        End If

Form1.Show
Unload Form2
End Sub

Private Sub Image4_Click()
End
End Sub

Private Sub Shop_false()
For Number = 0 To Shape4.UBound
    Shape4(Number).Visible = False
Next
For Number1 = 0 To Shape6.UBound
    Shape6(Number1).Visible = False
Next
For Number2 = 0 To Label2.UBound
    Label2(Number2).Visible = False
Next
For Number3 = 0 To Label3.UBound
    Label3(Number3).Visible = False
Next
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Label1.Visible = False
End Sub


Private Sub Label5_Click()
Open (App.Path & "\Data\Save1.for") For Output As #1
    Print #1, CannonHealth
    Print #1, CannonMax
    Print #1, CannonLevel
    Print #1, HouseHealth
    Print #1, HouseMax
    Print #1, HouseLevel
    Print #1, Gold
    Print #1, TotalScore
    Print #1, Levels
Close #1
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
    Label5.Caption = "Save1: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
Close #1
End Sub

Private Sub Label6_Click()
Open (App.Path & "\Data\Save2.for") For Output As #1
    Print #1, CannonHealth
    Print #1, CannonMax
    Print #1, CannonLevel
    Print #1, HouseHealth
    Print #1, HouseMax
    Print #1, HouseLevel
    Print #1, Gold
    Print #1, TotalScore
    Print #1, Levels
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
    Label6.Caption = "Save2: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
Close #1
End Sub

Private Sub Label7_Click()
Open (App.Path & "\Data\Save3.for") For Output As #1
    Print #1, CannonHealth
    Print #1, CannonMax
    Print #1, CannonLevel
    Print #1, HouseHealth
    Print #1, HouseMax
    Print #1, HouseLevel
    Print #1, Gold
    Print #1, TotalScore
    Print #1, Levels
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
    Label7.Caption = "Save3: Level: " & OLevels & " Gold: " & OGold & " Score: " & OTotalScore
Close #1
End Sub

Private Sub Save_false()
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
End Sub
