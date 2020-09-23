VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Circle Progress Bar"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton MeterBut 
      Caption         =   "One Down"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton MeterBut 
      Caption         =   "Get Value"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton MeterBut 
      Caption         =   "Set To"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton MeterBut 
      Caption         =   "One Up"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox SetBox 
      Height          =   285
      Left            =   1140
      TabIndex        =   4
      Text            =   "0"
      Top             =   2325
      Width           =   855
   End
   Begin VB.TextBox GetBox 
      Height          =   285
      Left            =   1140
      TabIndex        =   2
      Text            =   "0"
      Top             =   2670
      Width           =   855
   End
   Begin VB.PictureBox MeterBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      DrawWidth       =   2
      Height          =   1695
      Left            =   120
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.Label MeterPos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape MeterShape 
         Height          =   1350
         Left            =   240
         Shape           =   3  'Circle
         Top             =   120
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin VB.Label TitleLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CIRCLE PROGRESS BAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Index           =   0
      Left            =   2160
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label MyEmail 
      Alignment       =   2  'Center
      Caption         =   "dorejj@hotmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label InfoLabel 
      Alignment       =   2  'Center
      Caption         =   "Got a Suggestion, Comment, or a New Option you want me to add, please feel free to contact me at:"
      Height          =   1455
      Left            =   2160
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label TitleLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CIRCLE PROGRESS BAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   2190
      TabIndex        =   11
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LoadMeter Form1, RGB(0, 0, 128), RGB(255, 255, 255)
End Sub

Private Sub MeterBut_Click(Index As Integer)
Select Case Index
Case 0
    SetMeter -1, Form1
Case 1
    SetMeter -2, Form1
Case 2
    SetMeter SetBox.Text, Form1
Case 3
    GetMeter Form1
    GetBox.Text = MeterValue
End Select
End Sub
