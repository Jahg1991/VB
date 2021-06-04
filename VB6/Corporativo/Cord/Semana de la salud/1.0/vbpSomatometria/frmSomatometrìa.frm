VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Somatometrìa"
   ClientHeight    =   9750
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
   Icon            =   "frmSomatometrìa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   4800
      Picture         =   "frmSomatometrìa.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   3000
      Picture         =   "frmSomatometrìa.frx":3D4B
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   8
      Left            =   2880
      TabIndex        =   26
      Top             =   8040
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   360
      Index           =   7
      Left            =   3360
      TabIndex        =   25
      Top             =   7080
      Width           =   5895
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check2"
      Height          =   270
      Index           =   1
      Left            =   2880
      TabIndex        =   24
      Top             =   7080
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check2"
      Height          =   270
      Index           =   0
      Left            =   9000
      TabIndex        =   23
      Top             =   6120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   6
      Left            =   2880
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   5
      Left            =   8040
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   3240
      Width           =   6375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check1"
      Height          =   270
      Index           =   0
      Left            =   9000
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      Height          =   390
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   12648447
      Format          =   16842753
      CurrentDate     =   42553
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   14
      Left            =   8880
      Picture         =   "frmSomatometrìa.frx":A25D
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   13
      Left            =   8880
      Picture         =   "frmSomatometrìa.frx":A7E4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Observaciones"
      Height          =   375
      Index           =   12
      Left            =   960
      TabIndex        =   22
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Otras vacunas"
      Height          =   375
      Index           =   11
      Left            =   960
      TabIndex        =   21
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vacuna toxoide"
      Height          =   375
      Index           =   10
      Left            =   7080
      TabIndex        =   20
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tensiòn arterial"
      Height          =   375
      Index           =   9
      Left            =   960
      TabIndex        =   19
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Talla"
      Height          =   375
      Index           =   8
      Left            =   7320
      TabIndex        =   18
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Peso"
      Height          =   375
      Index           =   7
      Left            =   960
      TabIndex        =   17
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Familiar que labora en la empresa"
      Height          =   375
      Index           =   6
      Left            =   960
      TabIndex        =   16
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Àrea de trabajo"
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   15
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trabajor de la empresa"
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   14
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gènero"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edad"
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha de nacimiento"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   12
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":AD6B
      Top             =   7800
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   11
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":BB8C
      Top             =   6840
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   10
      Left            =   6240
      Picture         =   "frmSomatometrìa.frx":C3DA
      Top             =   5880
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   9
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":11C41
      Top             =   5880
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   8
      Left            =   6480
      Picture         =   "frmSomatometrìa.frx":126C6
      Top             =   4920
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   7
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":12DB1
      Top             =   4920
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   6
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":138A4
      Top             =   3960
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   5
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":14353
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   4
      Left            =   5400
      Picture         =   "frmSomatometrìa.frx":14C80
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   3
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":1574D
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   2
      Left            =   6480
      Picture         =   "frmSomatometrìa.frx":161F0
      Top             =   1080
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":16BD6
      Top             =   1080
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   120
      Picture         =   "frmSomatometrìa.frx":1757C
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Check1(0).Value = 0 Then
                Text1(2).Enabled = False
                Text1(2).Text = ""
                Image1(14).Visible = True
            Else
                Text1(2).Enabled = True
                Image1(14).Visible = False
            End If
    End Select
End Sub

Private Sub Check2_Click(Index As Integer)
    Select Case Index
        Case 1
            If Check2(1).Value = 0 Then
                Text1(7).Enabled = False
                Text1(7).Text = ""
            Else
                Text1(7).Enabled = True
            End If
    End Select
End Sub

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
            Text1(5).Text = ""
            Text1(6).Text = ""
            Text1(7).Text = ""
            Text1(8).Text = ""
            Check1(0).Value = 0
            Check2(0).Value = 0
            Check2(1).Value = 0
            DTPicker1.Value = Date
            Combo1(0) = "Femenino"
            Text1(2).Enabled = False
            Text1(7).Enabled = False
    End Select
End Sub

Private Sub DTPicker1_Change()
    On Error Resume Next
    Dim Fecha_Nacimiento As Date
    Dim Años As Variant
    Fecha_Nacimiento = DTPicker1.Value
    If IsNull(Fecha_Nacimiento) Then
        Calcular_Edad = 0
    End If
    Años = DateDiff("yyyy", Fecha_Nacimiento, Now)
    If Date < DateSerial(Year(Now), Month(Fecha_Nacimiento), Day(Fecha_Nacimiento)) Then
        Años = Años - 1
    End If
    Calcular_Edad = CInt(Años)
    Text1(1).Text = Años
End Sub

Private Sub DTPicker1_CloseUp()
    On Error Resume Next
    Dim Fecha_Nacimiento As Date
    Dim Años As Variant
    Fecha_Nacimiento = DTPicker1.Value
    If IsNull(Fecha_Nacimiento) Then
        Calcular_Edad = 0
    End If
    Años = DateDiff("yyyy", Fecha_Nacimiento, Now)
    If Date < DateSerial(Year(Now), Month(Fecha_Nacimiento), Day(Fecha_Nacimiento)) Then
        Años = Años - 1
    End If
    Calcular_Edad = CInt(Años)
    Text1(1).Text = Años
End Sub

Private Sub Form_Load()
    Combo1(0).AddItem ("Femenino")
    Combo1(0).AddItem ("Masculino")
    Combo1(0) = "Femenino"
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 13
            Form2.Show
            Form1.Enabled = False
        Case 14
            Form3.Show
            Form1.Enabled = False
    End Select
End Sub


