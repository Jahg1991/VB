VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   615
      Index           =   1
      Left            =   2400
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Index           =   0
      Left            =   600
      Picture         =   "frmLogin.frx":0B01
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   390
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Top             =   240
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   4455
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   390
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Usuarios"
         Enabled         =   0   'False
         Height          =   270
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Impresion"
         Enabled         =   0   'False
         Height          =   270
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tuberculosis"
         Enabled         =   0   'False
         Height          =   270
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Somatometria"
         Enabled         =   0   'False
         Height          =   270
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Optometria"
         Enabled         =   0   'False
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Nutricion"
         Enabled         =   0   'False
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mamografia"
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Laboratorio"
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Docm"
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Doccu"
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dental"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cardiologia"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Audiometria"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "frmLogin.frx":1667
      Stretch         =   -1  'True
      Top             =   960
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   120
      Picture         =   "frmLogin.frx":3A18
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
