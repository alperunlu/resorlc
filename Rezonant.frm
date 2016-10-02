VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rezonans Hesabý"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   Icon            =   "Rezonant.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HESAPLA"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   "C"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "F"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "L"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L, F, C, X1, C1, pi As Double
Private Sub Command1_Click()
If Option1.Value = True Then
pi = 3.14159265358979
L = 1 / (((2 * pi * F) ^ 2) * C)
F = Text2.Text
C = Text3.Text
Text1.Text = FormatNumber(L, 15, vbUseDefault)
Else
pi = 3.14159265358979
L = Text1.Text
F = Text2.Text
C = 1 / (((2 * pi * F) ^ 2) * L)
Text3.Text = FormatNumber(C, 15, vbUseDefault)
End If
End Sub
Private Sub Option1_click()
Text1.Enabled = False
Text3.Enabled = True
End Sub
Private Sub Option2_click()
Text3.Enabled = False
Text1.Enabled = True
End Sub

