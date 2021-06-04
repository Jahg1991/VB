VERSION 5.00
Begin VB.Form frmMenuInicial 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   12
      Left            =   4440
      Picture         =   "frmMenuInicial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   11
      Left            =   8760
      Picture         =   "frmMenuInicial.frx":139C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   10
      Left            =   5880
      Picture         =   "frmMenuInicial.frx":A006
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   9
      Left            =   3000
      Picture         =   "frmMenuInicial.frx":C373
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   8
      Left            =   120
      Picture         =   "frmMenuInicial.frx":14FD2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   7
      Left            =   8760
      Picture         =   "frmMenuInicial.frx":1D68F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   6
      Left            =   5880
      Picture         =   "frmMenuInicial.frx":1EE27
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   5
      Left            =   3000
      Picture         =   "frmMenuInicial.frx":20101
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   4
      Left            =   120
      Picture         =   "frmMenuInicial.frx":2181D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   3
      Left            =   8760
      Picture         =   "frmMenuInicial.frx":23C83
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   2
      Left            =   5880
      Picture         =   "frmMenuInicial.frx":252E2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   1
      Left            =   3000
      Picture         =   "frmMenuInicial.frx":27126
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   975
      Index           =   0
      Left            =   120
      Picture         =   "frmMenuInicial.frx":283EF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMenuInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Audiometria.exe"), vbNormalFocus
        Case 1
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Cardiologia.exe"), vbNormalFocus
        Case 2
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Dental.exe"), vbNormalFocus
        Case 3
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Doccu.exe"), vbNormalFocus
        Case 4
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Docm.exe"), vbNormalFocus
        Case 5
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Laboratorio.exe"), vbNormalFocus
        Case 6
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Mamografia.exe"), vbNormalFocus
        Case 7
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Nutricion.exe"), vbNormalFocus
        Case 8
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Opometria.exe"), vbNormalFocus
        Case 9
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Somatometria.exe"), vbNormalFocus
        Case 10
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Tuberculosis.exe"), vbNormalFocus
        Case 11
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Impresion.exe"), vbNormalFocus
        Case 12
            Shell ("C:\JAHG Software\Semana de la Salud 2016\Utilerias\Audiometria.exe"), vbNormalFocus0
    End Select
End Sub
