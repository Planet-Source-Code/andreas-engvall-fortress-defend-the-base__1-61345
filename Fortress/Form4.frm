VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form4"
   ScaleHeight     =   5625
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Main Menu"
      Height          =   315
      Left            =   4950
      TabIndex        =   0
      Top             =   5300
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "Form4.frx":0000
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form3
Form3.Show
Unload Form4
End Sub
