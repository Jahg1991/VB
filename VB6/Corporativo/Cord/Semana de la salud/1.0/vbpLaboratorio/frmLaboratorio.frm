VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorio"
   ClientHeight    =   5925
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
   Icon            =   "frmLaboratorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   4680
      Picture         =   "frmLaboratorio.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   2880
      Picture         =   "frmLaboratorio.frx":3D4B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   4
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   4
      Top             =   4200
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   2
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1320
      Width           =   975
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
   Begin VB.Label Label2 
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   13
      Left            =   8880
      Picture         =   "frmLaboratorio.frx":A25D
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Observaciones"
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Glucosa"
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trigliceridos"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Colesterol"
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   6
      Left            =   120
      Picture         =   "frmLaboratorio.frx":A7E4
      Top             =   3960
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   5
      Left            =   120
      Picture         =   "frmLaboratorio.frx":B605
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   3
      Left            =   120
      Picture         =   "frmLaboratorio.frx":BE8E
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   120
      Picture         =   "frmLaboratorio.frx":C94D
      Top             =   1080
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   120
      Picture         =   "frmLaboratorio.frx":D162
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
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
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
