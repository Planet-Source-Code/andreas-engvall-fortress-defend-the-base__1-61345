VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Fortress - Fight"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6900
      Top             =   150
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2550
      Top             =   300
   End
   Begin VB.Timer Timer3 
      Interval        =   1500
      Left            =   1950
      Top             =   300
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   1050
      Top             =   300
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   150
      Top             =   300
   End
   Begin VB.Shape Shape5 
      Height          =   1815
      Left            =   3750
      Top             =   1500
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Total Score:"
      Height          =   315
      Index           =   5
      Left            =   3750
      TabIndex        =   9
      Top             =   3000
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Kings Killed:"
      Height          =   315
      Index           =   4
      Left            =   3750
      TabIndex        =   8
      Top             =   2700
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Sorcerers killed:"
      Height          =   315
      Index           =   3
      Left            =   3750
      TabIndex        =   7
      Top             =   2400
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   4350
      Picture         =   "Form1.frx":0000
      Top             =   3300
      Width           =   900
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Princess Killed:"
      Height          =   315
      Index           =   0
      Left            =   3750
      TabIndex        =   6
      Top             =   1500
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Dwarfs Killed: "
      Height          =   315
      Index           =   1
      Left            =   3750
      TabIndex        =   5
      Top             =   1800
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Devils Killed:"
      Height          =   315
      Index           =   2
      Left            =   3750
      TabIndex        =   4
      Top             =   2100
      Width           =   2265
      Visible         =   0   'False
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   600
      Top             =   0
      Width           =   4065
      Visible         =   0   'False
   End
   Begin VB.Image Image16 
      Height          =   480
      Index           =   4
      Left            =   7500
      Picture         =   "Form1.frx":0618
      Top             =   4050
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.Image Image16 
      Height          =   480
      Index           =   3
      Left            =   7050
      Picture         =   "Form1.frx":0A4C
      Top             =   3900
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.Image Image16 
      Height          =   480
      Index           =   2
      Left            =   6600
      Picture         =   "Form1.frx":0E80
      Top             =   3600
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.Image Image16 
      Height          =   480
      Index           =   1
      Left            =   6150
      Picture         =   "Form1.frx":12B4
      Top             =   3300
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.Image Image16 
      Height          =   480
      Index           =   0
      Left            =   5700
      Picture         =   "Form1.frx":16E8
      Top             =   3000
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.Image Image15 
      Height          =   3300
      Left            =   0
      Picture         =   "Form1.frx":1B1C
      Top             =   0
      Width           =   5640
      Visible         =   0   'False
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   2100
      Top             =   2250
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   1200
      Top             =   2250
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   2100
      Top             =   3450
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   1050
      Top             =   1200
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   1950
      Top             =   1200
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Image Image14 
      Height          =   990
      Left            =   2100
      Picture         =   "Form1.frx":53EB
      Top             =   2250
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image Image13 
      Height          =   930
      Left            =   1050
      Picture         =   "Form1.frx":59D0
      Top             =   2400
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.Image Image12 
      Height          =   720
      Left            =   1950
      Picture         =   "Form1.frx":604F
      Top             =   3450
      Width           =   840
      Visible         =   0   'False
   End
   Begin VB.Image Image11 
      Height          =   930
      Left            =   1050
      Picture         =   "Form1.frx":65F1
      Top             =   1350
      Width           =   810
      Visible         =   0   'False
   End
   Begin VB.Image Image10 
      Height          =   990
      Left            =   1950
      Picture         =   "Form1.frx":6BF9
      Top             =   1350
      Width           =   720
      Visible         =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level:"
      Height          =   315
      Left            =   5550
      TabIndex        =   3
      Top             =   150
      Width           =   1215
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   7650
      Top             =   3450
      Width           =   1065
      Visible         =   0   'False
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   7650
      Top             =   1800
      Width           =   915
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   4
      Left            =   8850
      Top             =   1500
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   3
      Left            =   8700
      Top             =   1500
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   2
      Left            =   8550
      Top             =   1500
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   1
      Left            =   8400
      Top             =   1500
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   0
      Left            =   8250
      Top             =   1500
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Image8 
      Height          =   1290
      Left            =   7950
      Picture         =   "Form1.frx":7242
      Top             =   1650
      Width           =   1260
      Visible         =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Left: 40"
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   2400
      Picture         =   "Form1.frx":7CB3
      Top             =   150
      Width           =   300
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "        x0"
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Image Image6 
      Height          =   3450
      Left            =   5250
      Picture         =   "Form1.frx":8088
      Top             =   1350
      Width           =   465
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   4
      Left            =   9000
      Top             =   3150
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   3
      Left            =   8850
      Top             =   3150
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   2
      Left            =   8700
      Top             =   3150
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   1
      Left            =   8550
      Top             =   3150
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Index           =   0
      Left            =   8400
      Top             =   3150
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Image5 
      Height          =   1560
      Left            =   7950
      Picture         =   "Form1.frx":8A99
      Top             =   3300
      Width           =   1470
      Visible         =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GET READY"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   3150
      TabIndex        =   0
      Top             =   2100
      Width           =   3870
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   150
      Top             =   1650
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   150
      Top             =   2550
      Width           =   765
      Visible         =   0   'False
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   150
      Top             =   3750
      Width           =   760
      Visible         =   0   'False
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   0
      Picture         =   "Form1.frx":9AB3
      Top             =   3900
      Width           =   960
      Visible         =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   990
      Left            =   150
      Picture         =   "Form1.frx":A19D
      Top             =   2700
      Width           =   720
      Visible         =   0   'False
   End
   Begin VB.Image Image3 
      Height          =   780
      Left            =   150
      Picture         =   "Form1.frx":A7D6
      Top             =   1800
      Width           =   870
      Visible         =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   5070
      Left            =   0
      Picture         =   "Form1.frx":AE35
      Top             =   -150
      Width           =   9540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrinNr As Integer
Dim PrinHealth As Integer
Dim PrinMax As Integer
Dim PrinKilled As Integer
Dim DwarfNr As Integer
Dim DwarfHealth As Integer
Dim DwarfMax As Integer
Dim DwarfKilled As Integer
Dim DevNr As Integer
Dim DevHealth As Integer
Dim DevMax As Integer
Dim DevKilled As Integer
Dim NumberPic As Integer
Dim AmaHealth As Integer
Dim AmaMax As Integer
Dim AmaKilled As Integer
Dim ThiHealth As Integer
Dim ThiMax As Integer
Dim ThiKilled As Integer
Dim PixHealth As Integer
Dim PixMax As Integer
Dim PixKilled As Integer
Dim SorHealth As Integer
Dim SorMax As Integer
Dim SorKilled As Integer
Dim KingHealth As Integer
Dim KingMax As Integer
Dim KingKilled As Integer
Dim FireBreath As Integer
Dim EmperorHealth As Integer
Dim EmperorMax As Integer

Dim Score As Integer

Dim PlayHit
Dim PlayDiss


Private Sub Form_Load()
EmperorMax = 5000
EmperorHealth = EmperorMax
Shape14.Width = EmperorHealth
FireBreath = 0
Levels = 1
NumberPic = 1
Label5.Caption = "Levels: " & Levels
HouseLevel = 1
HouseMax = HouseLevel * 550
HouseHealth = HouseMax
Shape8.Width = HouseHealth
CannonLevel = 1
Damage = CannonLevel * 30
TimeLeft = 40
Gold = 0
PrinNr = 1
PrinMax = 100
PrinHealth = PrinMax
DwarfNr = 1
DwarfMax = 150
DwarfHealth = DwarfMax
DevNr = 1
DevMax = 120
DevHealth = DevMax
AmaMax = 160
AmaHealth = AmaMax
ThiMax = 150
ThiHealth = ThiMax
PixMax = 170
PixHealth = PixMax
SorMax = 180
SorHealth = SorMax
KingMax = 190
KingHealth = KingMax
CannonMax = 500
CannonHealth = CannonMax
Shape7.Width = CannonHealth
Label2.Caption = "       x" & Gold
End Sub

Private Sub Image10_Click()
If Timer4.Enabled = False Then
Else
   PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    AmaHealth = AmaHealth - Damage
    If AmaHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image10.Visible = False
        Timer2.Enabled = True
        Shape9.Visible = False
        Gold = Gold + 70 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        AmaKilled = AmaKilled + 1
    Else
        Shape9.Width = AmaHealth * 6
    End If
End If
End Sub

Private Sub Image11_Click()
If Timer4.Enabled = False Then
Else
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    ThiHealth = ThiHealth - Damage
    If ThiHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image11.Visible = False
        Timer2.Enabled = True
        Shape10.Visible = False
        Gold = Gold + 65 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        ThiKilled = ThiKilled + 1
    Else
        Shape10.Width = ThiHealth * 6
    End If
End If
End Sub

Private Sub Image12_Click()
If Timer4.Enabled = False Then
Else
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    PixHealth = PixHealth - Damage
    If PixHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image12.Visible = False
        Timer2.Enabled = True
        Shape11.Visible = False
        Gold = Gold + 55 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        PixKilled = PixKilled + 1
    Else
        Shape11.Width = PixHealth * 6
    End If
End If
End Sub

Private Sub Image13_Click()
If Timer4.Enabled = False Then
Else
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    SorHealth = SorHealth - Damage
    If SorHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image13.Visible = False
        Timer2.Enabled = True
        Shape12.Visible = False
        Gold = Gold + 70 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        SorKilled = SorKilled + 1
    Else
        Shape12.Width = SorHealth * 6
    End If
End If
End Sub

Private Sub Image14_Click()
If Timer4.Enabled = False Then
Else
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    KingHealth = KingHealth - Damage
    If KingHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image14.Visible = False
        Timer2.Enabled = True
        Shape13.Visible = False
        Gold = Gold + 80 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        KingKilled = KingKilled + 1
    Else
        Shape13.Width = KingHealth * 6
    End If
End If
End Sub

Private Sub Image15_Click()
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    EmperorHealth = EmperorHealth - Damage
    If EmperorHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image15.Visible = False
        Shape14.Visible = False
        Load Form4
        Form4.Show
        Unload Form1
        TimerEnabled_False
    Else
        Shape14.Width = EmperorHealth
    End If
End Sub



Private Sub Image2_Click()
If Timer4.Enabled = False Then
Else
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    PrinHealth = PrinHealth - Damage
    If PrinHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image2.Visible = False
        Timer2.Enabled = True
        Shape2.Visible = False
        Gold = Gold + 41 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        PrinKilled = PrinKilled + 1
    Else
        Shape2.Width = PrinHealth * 6
    End If
End If
End Sub

Private Sub Image3_Click()
If Timer4.Enabled = False Then
Else
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    DwarfHealth = DwarfHealth - Damage
    If DwarfHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\Sound\disapear.wav", SND_ASYNC&)
        Image3.Visible = False
        Timer2.Enabled = True
        Shape3.Visible = False
        Gold = Gold + 47 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        DwarfKilled = DwarfKilled + 1
    Else
        Shape3.Width = DwarfHealth * 6
    End If
End If
End Sub

Private Sub Image4_Click()
If Timer4.Enabled = False Then
Else

    DevHealth = DevHealth - Damage
    If DevHealth <= 0 Then
        PlayDiss = sndPlaySound(App.Path & "\\Sound\disapear.wav", SND_ASYNC&)
        Image4.Visible = False
        Timer2.Enabled = True
        Shape1.Visible = False
        Gold = Gold + 52 + (Levels * 4)
        Label2.Caption = "       x" & Gold
        DevKilled = DevKilled + 1
    Else
        Shape1.Width = DevHealth * 6
PlayHit = sndPlaySound(App.Path & "\Sound\hit.wav", SND_ASYNC&)
    End If
End If
End Sub





Private Sub Image9_Click()
Load Form2
Form2.Show
Form1.Hide
Form1.Enabled = False
End Sub





Private Sub Timer1_Timer()
If HouseHealth <= 0 Then
    Dim Svar
    Svar = MsgBox("You've lost, press OK to return to main menu.", vbOKOnly, "You've lost!")
    If Svar = 1 Then
        Load Form3
        Form3.Show
        Timer1.Enabled = False
        Unload Form1
    End If
End If
If Levels <= 5 Then
    PrinNr = PrinNr + 1
    Image2.Picture = LoadPicture(App.Path & "\Pics\Princess" & PrinNr & ".gif")
    If PrinNr = 6 Then PrinNr = 1
    If Image2.Visible = True Then
        If Image2.Left < 7050 Then
            Image2.Left = Image2.Left + 90 + (Levels * 10)
            Shape2.Left = Image2.Left + 70
            Shape2.Top = Image2.Top - 100
        Else
            If Image2.Top < 2400 Then
                If CannonHealth <= 1 Then
                    Image8.Visible = False
                    Shape7.Visible = False
                    For Number = 0 To Shape6.UBound
                        Shape6(Number).Visible = False
                    Next
                    Damage = 15
    
                Else
                    Shape7.Width = CannonHealth
                    CannonHealth = CannonHealth - 5 - Levels
                End If
            Else
                If HouseHealth <= 1 Then
                    Image5.Visible = False
                    Shape8.Visible = False
                    For Numbe1 = 0 To Shape4.UBound
                        Shape4(Numbe1).Visible = False
                    Next
                        Shape8.Width = 0
                Else
                    Shape8.Width = HouseHealth
                    HouseHealth = HouseHealth - 5 + Levels
                End If
            End If
        End If
    End If
    
    DwarfNr = DwarfNr + 1
    Image3.Picture = LoadPicture(App.Path & "\Pics\Dwarf" & DwarfNr & ".gif")
    If DwarfNr = 6 Then DwarfNr = 1
    If Image3.Visible = True Then
    If Image3.Left < 7050 Then
        Image3.Left = Image3.Left + 50 + (Levels * 10)
        Shape3.Left = Image3.Left + 70
        Shape3.Top = Image3.Top - 100
    Else
        If Image3.Top < 2400 Then
            If CannonHealth <= 1 Then
                Image8.Visible = False
                Shape7.Visible = False
                For Number = 0 To Shape6.UBound
                    Shape6(Number).Visible = False
                Next
                Damage = 15
            Else
                Shape7.Width = CannonHealth
                CannonHealth = CannonHealth - 5 - Levels
            End If
        Else
            If HouseHealth <= 1 Then
                Image5.Visible = False
                Shape8.Visible = False
                For Numbe1 = 0 To Shape4.UBound
                    Shape4(Numbe1).Visible = False
                Next
                Shape8.Width = 0
            Else
                Shape8.Width = HouseHealth
                HouseHealth = HouseHealth - 5 - Levels
            End If
        End If
    End If
    End If
    
    DevNr = DevNr + 1
    Image4.Picture = LoadPicture(App.Path & "\Pics\Devil" & DevNr & ".gif")
    If DevNr = 6 Then DevNr = 1
    If Image4.Visible = True Then
    If Image4.Left < 7050 Then
        Image4.Left = Image4.Left + 110 + (Levels * 10)
        Shape1.Left = Image4.Left + 70
        Shape1.Top = Image4.Top - 100
    Else
        If Image4.Top < 2400 Then
            If CannonHealth <= 1 Then
                Image8.Visible = False
                Shape7.Visible = False
                For Number = 0 To Shape6.UBound
                    Shape6(Number).Visible = False
                Next
                Damage = 15
            Else
                Shape7.Width = CannonHealth
                CannonHealth = CannonHealth - 5 - Levels
            End If
        Else
            If HouseHealth <= 1 Then
                Image5.Visible = False
                Shape8.Visible = False
                For Numbe1 = 0 To Shape4.UBound
                    Shape4(Numbe1).Visible = False
                Next
                Shape8.Width = 0
            Else
                Shape8.Width = HouseHealth
                HouseHealth = HouseHealth - 5 + Levels
            End If
        End If
    End If
    End If
    
ElseIf Levels > 5 Then
    If Levels < 13 Then
        NumberPic = NumberPic + 1
        If NumberPic = 6 Then NumberPic = 1
        
        Image10.Picture = LoadPicture(App.Path & "\Pics\Amazoness" & NumberPic & ".gif")
        If Image10.Visible = True Then
        If Image10.Left < 7050 Then
            Image10.Left = Image10.Left + 150 + (Levels * 10)
            Shape9.Left = Image10.Left + 70
            Shape9.Top = Image10.Top - 100
        Else
            If Image10.Top < 2400 Then
                If CannonHealth <= 1 Then
                    Image8.Visible = False
                    Shape7.Visible = False
                    For Number = 0 To Shape6.UBound
                        Shape6(Number).Visible = False
                    Next
                    Damage = 15
                Else
                    Shape7.Width = CannonHealth
                    CannonHealth = CannonHealth - 10 - Levels
                End If
            Else
                If HouseHealth <= 1 Then
                    Image5.Visible = False
                    Shape8.Visible = False
                    For Numbe1 = 0 To Shape4.UBound
                        Shape4(Numbe1).Visible = False
                    Next
                    Shape8.Width = 0
                Else
                    Shape8.Width = HouseHealth
                    HouseHealth = HouseHealth - 10 - Levels
                End If
            End If
        End If
        End If
        
        Image11.Picture = LoadPicture(App.Path & "\Pics\Thief" & NumberPic & ".gif")
        If Image11.Visible = True Then
        If Image11.Left < 7050 Then
            Image11.Left = Image11.Left + 170 + (Levels * 10)
            Shape10.Left = Image11.Left + 70
            Shape10.Top = Image11.Top - 100
        Else
            If Image11.Top < 2400 Then
                If CannonHealth <= 1 Then
                    Image8.Visible = False
                    Shape7.Visible = False
                    For Number = 0 To Shape6.UBound
                        Shape6(Number).Visible = False
                    Next
                    Damage = 15
                Else
                    Shape7.Width = CannonHealth
                    CannonHealth = CannonHealth - 7 - Levels
                End If
            Else
                If HouseHealth <= 1 Then
                    Image5.Visible = False
                    Shape8.Visible = False
                    For Numbe1 = 0 To Shape4.UBound
                        Shape4(Numbe1).Visible = False
                    Next
                    Shape8.Width = 0
                Else
                    Shape8.Width = HouseHealth
                    HouseHealth = HouseHealth - 7 - Levels
                End If
            End If
        End If
        End If
        
        Image12.Picture = LoadPicture(App.Path & "\Pics\Pixie" & NumberPic & ".gif")
        If Image12.Visible = True Then
        If Image12.Left < 7050 Then
            Image12.Left = Image12.Left + 140 + (Levels * 10)
            Shape11.Left = Image12.Left + 70
            Shape11.Top = Image12.Top - 100
        Else
            If Image12.Top < 2400 Then
                If CannonHealth <= 1 Then
                    Image8.Visible = False
                    Shape7.Visible = False
                    For Number = 0 To Shape6.UBound
                        Shape6(Number).Visible = False
                    Next
                    Damage = 15
                Else
                    Shape7.Width = CannonHealth
                    CannonHealth = CannonHealth - 8 - Levels
                End If
            Else
                If HouseHealth <= 1 Then
                    Image5.Visible = False
                    Shape8.Visible = False
                    For Numbe1 = 0 To Shape4.UBound
                        Shape4(Numbe1).Visible = False
                    Next
                    Shape8.Width = 0
                Else
                    Shape8.Width = HouseHealth
                    HouseHealth = HouseHealth - 8 - Levels
                End If
            End If
        End If
        End If
        
        Image13.Picture = LoadPicture(App.Path & "\Pics\Sorceress" & NumberPic & ".gif")
        If Image13.Visible = True Then
        If Image13.Left < 7050 Then
            Image13.Left = Image13.Left + 120 + (Levels * 10)
            Shape12.Left = Image13.Left + 70
            Shape12.Top = Image13.Top - 100
        Else
            If Image13.Top < 2400 Then
                If CannonHealth <= 1 Then
                    Image8.Visible = False
                    Shape7.Visible = False
                    For Number = 0 To Shape6.UBound
                        Shape6(Number).Visible = False
                    Next
                    Damage = 15
                Else
                    Shape7.Width = CannonHealth
                    CannonHealth = CannonHealth - 11 - Levels
                End If
            Else
                If HouseHealth <= 1 Then
                    Image5.Visible = False
                    Shape8.Visible = False
                    For Numbe1 = 0 To Shape4.UBound
                        Shape4(Numbe1).Visible = False
                    Next
                    Shape8.Width = 0
                Else
                    Shape8.Width = HouseHealth
                    HouseHealth = HouseHealth - 11 - Levels
                End If
            End If
        End If
        End If
        
        Image14.Picture = LoadPicture(App.Path & "\Pics\King" & NumberPic & ".gif")
        If Image14.Visible = True Then
        If Image14.Left < 7050 Then
            Image14.Left = Image14.Left + 100 + (Levels * 10)
            Shape13.Left = Image14.Left + 70
            Shape13.Top = Image14.Top - 100
        Else
            If Image14.Top < 2400 Then
                If CannonHealth <= 1 Then
                    Image8.Visible = False
                    Shape7.Visible = False
                    For Number = 0 To Shape6.UBound
                        Shape6(Number).Visible = False
                    Next
                    Damage = 15
                Else
                    Shape7.Width = CannonHealth
                    CannonHealth = CannonHealth - 12 - Levels
                End If
            Else
                If HouseHealth <= 1 Then
                    Image5.Visible = False
                    Shape8.Visible = False
                    For Numbe1 = 0 To Shape4.UBound
                        Shape4(Numbe1).Visible = False
                    Next
                    Shape8.Width = 0
                Else
                    Shape8.Width = HouseHealth
                    HouseHealth = HouseHealth - 12 - Levels
                End If
            End If
        End If
        End If
        
    End If
End If
End Sub

Private Sub Timer2_Timer()
If Levels <= 5 Then
    If Image2.Visible = False Then
        Dim Prin
        PrinHealth = PrinMax
        Shape2.Width = PrinHealth * 6
        
        Randomize   ' Initialize random-number generator.
        Prin = Int((3600 * Rnd) + 500)
        Image2.Left = -1200
        Image2.Top = Prin
        Image2.Visible = True
        
        Timer1.Enabled = True
        Shape2.Visible = True
        Shape2.Left = Image2.Left + 70
        Shape2.Top = Image2.Top - 130
        Timer2.Enabled = False
    End If
    If Image3.Visible = False Then
        Dim Dwarf
        DwarfHealth = DwarfMax
        Shape3.Width = DwarfHealth * 6
        
        Randomize   ' Initialize random-number generator.
        Dwarf = Int((3700 * Rnd) + 700)
        Image3.Left = -1000
        Image3.Top = Dwarf
        Image3.Visible = True
        
        Timer1.Enabled = True
        Shape3.Visible = True
        Shape3.Left = Image3.Left + 80
        Shape3.Top = Image3.Top - 100
        Timer2.Enabled = False
    End If
    If Image4.Visible = False Then
        Dim Devils
        DevHealth = DevMax
        Shape1.Width = DevHealth * 6
        
        Randomize   ' Initialize random-number generator.
        Devils = Int((3700 * Rnd) + 600)
        Image4.Left = -1000
        Image4.Top = Devils
        Image4.Visible = True
        
        Timer1.Enabled = True
        Shape1.Visible = True
        Shape1.Left = Image4.Left + 80
        Shape1.Top = Image4.Top - 100
        Timer2.Enabled = False
    End If
ElseIf Levels < 13 Then
    If Levels > 5 Then
        If Image10.Visible = False Then
            Dim Amazonessa
            AmaHealth = AmaMax
            Shape9.Width = AmaHealth * 6
            
            Randomize   ' Initialize random-number generator.
            Amazonessa = Int((3700 * Rnd) + 600)
            Image10.Left = -1000
            Image10.Top = Amazonessa
            Image10.Visible = True
            
            Timer1.Enabled = True
            Shape9.Visible = True
            Shape9.Left = Image10.Left + 80
            Shape9.Top = Image10.Top - 100
            Timer2.Enabled = False
        End If
        If Image11.Visible = False Then
            Dim Thiefa
            ThiHealth = ThiMax
            Shape10.Width = ThiHealth * 6
            
            Randomize   ' Initialize random-number generator.
            Thiefa = Int((3700 * Rnd) + 600)
            Image11.Left = -1000
            Image11.Top = Thiefa
            Image11.Visible = True
            
            Timer1.Enabled = True
            Shape10.Visible = True
            Shape10.Left = Image11.Left + 80
            Shape10.Top = Image11.Top - 100
            Timer2.Enabled = False
        End If
        If Image12.Visible = False Then
            Dim Pixies
            PixHealth = PixMax
            Shape11.Width = PixHealth * 6
            
            Randomize   ' Initialize random-number generator.
            Pixies = Int((3700 * Rnd) + 600)
            Image12.Left = -1000
            Image12.Top = Pixies
            Image12.Visible = True
            
            Timer1.Enabled = True
            Shape11.Visible = True
            Shape11.Left = Image12.Left + 80
            Shape11.Top = Image12.Top - 100
            Timer2.Enabled = False
        End If
        If Image13.Visible = False Then
            Dim Sorceressa
            SorHealth = SorMax
            Shape12.Width = SorHealth * 6
            
            Randomize   ' Initialize random-number generator.
            Sorceressa = Int((3700 * Rnd) + 600)
            Image13.Left = -1000
            Image13.Top = Sorceressa
            Image13.Visible = True
            
            Timer1.Enabled = True
            Shape12.Visible = True
            Shape12.Left = Image13.Left + 80
            Shape12.Top = Image13.Top - 100
            Timer2.Enabled = False
        End If
        If Image14.Visible = False Then
            Dim Kings
            KingHealth = KingMax
            Shape13.Width = KingHealth * 6
            
            Randomize   ' Initialize random-number generator.
            Kings = Int((3700 * Rnd) + 600)
            Image14.Left = -1000
            Image14.Top = Kings
            Image14.Visible = True
            
            Timer1.Enabled = True
            Shape13.Visible = True
            Shape13.Left = Image14.Left + 80
            Shape13.Top = Image14.Top - 100
            Timer2.Enabled = False
        End If
    End If
ElseIf Levels = 13 Then
    Timer4.Enabled = False
    Image15.Visible = True
    Timer5.Enabled = True

End If

End Sub

Private Sub Timer3_Timer()
If Levels < 6 Then
Image1.Picture = LoadPicture(App.Path & "\Pics\BG1.gif")
Else
Image1.Picture = LoadPicture(App.Path & "\Pics\BG2.gif")
End If
If Levels = 13 Then Shape14.Visible = True
Label1.Visible = False
Image5.Visible = True
Image8.Visible = True
Shape7.Visible = True
Shape8.Visible = True
Timer1.Enabled = True
For Number = 0 To Shape4.UBound
    Shape4(Number).Visible = True
Next
For Number1 = 0 To Shape6.UBound
    Shape6(Number1).Visible = True
Next

Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
TimeLeft = TimeLeft - 1
Label3.Caption = "Time Left: " & TimeLeft

If TimeLeft = 0 Then
    For Number = 0 To Label4.UBound
        Label4(Number).Visible = True
    Next
    Shape5.Visible = True
    If Levels < 6 Then
        Label4(0).Caption = "Princess: " & PrinKilled & "x32 = " & PrinKilled * 32
        Label4(1).Caption = "Dwarfs: " & DwarfKilled & "x36 = " & DwarfKilled * 36
        Label4(2).Caption = "Devils: " & DevKilled & "x40 = " & DevKilled * 40
        TotalScore = TotalScore + (PrinKilled * 32) + (DwarfKilled * 36) + (DevKilled * 40)
        Label4(3).Caption = "Total score: " & TotalScore
    ElseIf Levels >= 6 Then
        If Levels <= 13 Then
        Label4(0).Caption = "Amazons: " & AmaKilled & "x50 = " & AmaKilled * 50
        Label4(1).Caption = "Pixies: " & PixKilled & "x53 = " & PixKilled * 53
        Label4(2).Caption = "Thieves: " & ThiKilled & "x47 = " & ThiKilled * 47
        Label4(3).Caption = "Sorcerers: " & SorKilled & "x56 = " & SorKilled * 56
        Label4(4).Caption = "Kings: " & KingKilled & "x60 = " & KingKilled * 60
        TotalScore = TotalScore + (AmaKilled * 50) + (PixKilled * 53) + (ThiKilled * 47) + (SorKilled * 56) + (KingKilled * 60)
        Label4(5).Caption = "Totalscore: " & TotalScore
        End If

End If
    Image9.Visible = True
    PrinKilled = 0
    DwarfKilled = 0
    DevKilled = 0
    AmaKilled = 0
    PixKilled = 0
    ThiKilled = 0
    SorKilled = 0
    KingKilled = 0
TimerEnabled_False

End If
End Sub

Private Sub TimerEnabled_False()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
End Sub

Private Sub Timer5_Timer()
Image16(FireBreath).Visible = True
FireBreath = FireBreath + 1
If FireBreath = 5 Then
    FireBreath = 0
    For Number = 0 To Image16.UBound
        Image16(Number).Visible = False
    Next
    If Shape8.Width >= 280 Then
        HouseHealth = HouseHealth - 280
        Shape8.Width = HouseHealth
    Else
        Dim Svar
        Svar = MsgBox("You've lost, press OK to return to main menu.", vbOKOnly, "You've lost!")
        If Svar = 1 Then
        Load Form3
        Form3.Show
        Timer1.Enabled = False
        Unload Form1
    End If
End If
    End If


End Sub
