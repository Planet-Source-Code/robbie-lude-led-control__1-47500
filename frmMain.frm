VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LED Control"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   173
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   8
      Left            =   1170
      TabIndex        =   28
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   1170
      TabIndex        =   27
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   1410
      TabIndex        =   26
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   930
      TabIndex        =   25
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   1410
      TabIndex        =   24
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   930
      TabIndex        =   23
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   1170
      TabIndex        =   22
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   1170
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox c 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   1170
      TabIndex        =   20
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   19
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   18
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   17
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   16
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   15
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   12
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton o 
      BackColor       =   &H00404040&
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin Project1.xeroLED d 
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   661
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Which As Byte

Private Sub c_Click(Index As Integer)
    Select Case c(Index).Value
        Case 0
            d(Which).Value = d(Which).Value And 511 - (2 ^ Index)
        Case 1
            d(Which).Value = d(Which).Value Or (2 ^ Index)
    End Select
    l = d(Which).Value
End Sub

Private Sub o_Click(Index As Integer)
    Which = Index
    For i = 0 To 8
        c(i).Value = (d(Which).Value And (2 ^ i)) \ (2 ^ i)
    Next
    l = d(Which).Value
End Sub
