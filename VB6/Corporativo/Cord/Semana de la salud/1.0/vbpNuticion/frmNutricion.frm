VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nutrici�n"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNutricion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   4800
      Picture         =   "frmNutricion.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   3000
      Picture         =   "frmNutricion.frx":3D4B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   6495
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   270
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   270
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Observaciones"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   2
      Left            =   120
      Picture         =   "frmNutricion.frx":A25D
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   8640
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   13
      Left            =   8880
      Picture         =   "frmNutricion.frx":B07E
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consulta"
      Height          =   375
      Index           =   12
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pl�tica general"
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   12
      Left            =   120
      Picture         =   "frmNutricion.frx":B605
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   120
      Picture         =   "frmNutricion.frx":D66F
      Top             =   1080
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   120
      Picture         =   "frmNutricion.frx":1B13A
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Form4.Show
            Form1.Enabled = False
        Case 1
            Text1(0).Text = ""
            Text1(1).Text = ""
            Check1(0).Value = 0
            Check1(1).Value = 0
            Label2 = ""
    End Select
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 13
            Form2.Show
            Form1.Enabled = False
    End Select
End Sub
