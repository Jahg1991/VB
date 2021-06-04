VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5835
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.838
   ScaleMode       =   0  'User
   ScaleWidth      =   5478.749
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   5520
      TabIndex        =   35
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   2400
      TabIndex        =   33
      Top             =   1920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Archivo"
      TabPicture(0)   =   "frmLogin.frx":324A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(1)=   "Check1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Módulos"
      TabPicture(1)   =   "frmLogin.frx":3266
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(8)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(35)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(13)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(14)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Check2(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Check2(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Check2(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Check2(3)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Check2(4)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Check2(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Check2(6)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Check2(7)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Check2(8)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Check2(9)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Edición"
      TabPicture(2)   =   "frmLogin.frx":3282
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check3(2)"
      Tab(2).Control(1)=   "Check3(1)"
      Tab(2).Control(2)=   "Check3(0)"
      Tab(2).Control(3)=   "Label1(9)"
      Tab(2).Control(4)=   "Label1(1)"
      Tab(2).Control(5)=   "Label1(0)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Reportes"
      TabPicture(3)   =   "frmLogin.frx":329E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check4(1)"
      Tab(3).Control(1)=   "Check4(0)"
      Tab(3).Control(2)=   "Label1(11)"
      Tab(3).Control(3)=   "Label1(10)"
      Tab(3).ControlCount=   4
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   9
         Left            =   3120
         TabIndex        =   40
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   8
         Left            =   3120
         TabIndex        =   39
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   7
         Left            =   3120
         TabIndex        =   38
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   270
         Index           =   1
         Left            =   -73080
         TabIndex        =   37
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Index           =   1
         Left            =   5040
         TabIndex        =   34
         Top             =   -1440
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   270
         Index           =   0
         Left            =   -73080
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   270
         Index           =   2
         Left            =   -73200
         TabIndex        =   28
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   270
         Index           =   1
         Left            =   -73200
         TabIndex        =   27
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   270
         Index           =   0
         Left            =   -73200
         TabIndex        =   26
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   6
         Left            =   3120
         TabIndex        =   22
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   5
         Left            =   3120
         TabIndex        =   20
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   4
         Left            =   3120
         TabIndex        =   19
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   3
         Left            =   3120
         TabIndex        =   18
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   2
         Left            =   3120
         TabIndex        =   17
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   1
         Left            =   3120
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Index           =   0
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   270
         Left            =   -73080
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cardiología"
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   43
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tuberculosis"
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   42
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Audiometría"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   41
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estadisticas"
         Height          =   375
         Index           =   11
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Familiares"
         Height          =   375
         Index           =   10
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Exportar"
         Height          =   375
         Index           =   9
         Left            =   -74880
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Importar"
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Editar"
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Impresión"
         Height          =   375
         Index           =   35
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Optometría"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Salud de la mujer"
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   12
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nutrición"
         Height          =   375
         Index           =   6
         Left            =   1680
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dental"
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Laboratorio"
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Somatometría"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuarios"
         Height          =   375
         Index           =   2
         Left            =   -74520
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   3405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   3405
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   3
      Left            =   3960
      TabIndex        =   32
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   1920
      Width           =   2130
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   2130
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   540
      Width           =   1425
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload frmLogin
    Unload frmAbout
    Unload Form1
End Sub
Private Sub cmdOK_Click()
    On Error Resume Next
    If txtUserName = Text1(0).Text And txtPassword = Text1(2).Text Then
        MsgBox "Bienvenido(a) " & txtUserName, , "Inicio de sesión"
        Form1.Enabled = True
        Form1.Caption = "Semana de la salud - " & txtUserName
        Unload frmLogin
    Else
        MsgBox "El usuario o contraseña no son válidos. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        txtPassword.Text = ""
    End If
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RRsUsers
        If .State = 1 Then .Close
            .Open "Select * from USERS", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Form1.Usuario.Enabled = False
    Form1.Somatometria.Enabled = False
    Form1.Laboratorio.Enabled = False
    Form1.Dental.Enabled = False
    Form1.Nutricion.Enabled = False
    Form1.Salud_mujer.Enabled = False
    Form1.Optometría.Enabled = False
    Form1.Impresion.Enabled = False
    Form1.Editar_eliminar_registros.Enabled = False
    Form1.Importar_Somatometria.Enabled = False
    Form1.Exportar_informacion.Enabled = False
    Form1.Familiares.Enabled = False
    Form1.Estatisticas.Enabled = False
    Form1.Audiometria.Enabled = False
    Form1.Tuberculosis.Enabled = False
    Form1.Cardiologia.Enabled = False
    Set frmLogin.Text1(0).DataSource = RRsUsers
    Set frmLogin.Text1(2).DataSource = RRsUsers
    Set frmLogin.Check1.DataSource = RRsUsers
    Set frmLogin.Check2(0).DataSource = RRsUsers
    Set frmLogin.Check2(1).DataSource = RRsUsers
    Set frmLogin.Check2(2).DataSource = RRsUsers
    Set frmLogin.Check2(3).DataSource = RRsUsers
    Set frmLogin.Check2(4).DataSource = RRsUsers
    Set frmLogin.Check2(5).DataSource = RRsUsers
    Set frmLogin.Check2(6).DataSource = RRsUsers
    Set frmLogin.Check2(7).DataSource = RRsUsers
    Set frmLogin.Check2(8).DataSource = RRsUsers
    Set frmLogin.Check2(9).DataSource = RRsUsers
    Set frmLogin.Check3(0).DataSource = RRsUsers
    Set frmLogin.Check3(1).DataSource = RRsUsers
    Set frmLogin.Check3(2).DataSource = RRsUsers
    Set frmLogin.Check4(0).DataSource = RRsUsers
    Set frmLogin.Check4(1).DataSource = RRsUsers
    frmLogin.Text1(0).DataField = ("NOMBRE")
    frmLogin.Text1(2).DataField = ("PASS")
    frmLogin.Check1.DataField = ("USUARIOS")
    frmLogin.Check2(0).DataField = ("SOMATOMETRIA")
    frmLogin.Check2(1).DataField = ("LABORATORIO")
    frmLogin.Check2(2).DataField = ("DENTAL")
    frmLogin.Check2(3).DataField = ("NUTRICION")
    frmLogin.Check2(4).DataField = ("MUJER")
    frmLogin.Check2(5).DataField = ("OPTOMETRIA")
    frmLogin.Check2(6).DataField = ("IMPRESION")
    frmLogin.Check2(7).DataField = ("AUDIOMETRIA")
    frmLogin.Check2(8).DataField = ("TUBERCULOSIS")
    frmLogin.Check2(9).DataField = ("CARDIOLOGIA")
    frmLogin.Check3(0).DataField = ("EDITAR")
    frmLogin.Check3(1).DataField = ("IMPORTAR")
    frmLogin.Check3(2).DataField = ("EXPORTAR")
    frmLogin.Check4(0).DataField = ("FAMILIARES")
    frmLogin.Check4(1).DataField = ("ESTADISTICAS")
    If frmLogin.Check1.Value = 1 Then
        Form1.Usuario.Enabled = True
    Else
        Form1.Usuario.Enabled = False
    End If
    If frmLogin.Check2(0).Value = 1 Then
        Form1.Somatometria.Enabled = True
    Else
        Form1.Somatometria.Enabled = False
    End If
    If frmLogin.Check2(1).Value = 1 Then
        Form1.Laboratorio.Enabled = True
    Else
        Form1.Laboratorio.Enabled = False
    End If
    If frmLogin.Check2(2).Value = 1 Then
        Form1.Dental.Enabled = True
    Else
        Form1.Dental.Enabled = False
    End If
    If frmLogin.Check2(3).Value = 1 Then
        Form1.Nutricion.Enabled = True
    Else
        Form1.Nutricion.Enabled = False
    End If
    If frmLogin.Check2(4).Value = 1 Then
        Form1.Salud_mujer.Enabled = True
    Else
        Form1.Salud_mujer.Enabled = False
    End If
    If frmLogin.Check2(5).Value = 1 Then
        Form1.Optometría.Enabled = True
    Else
        Form1.Optometría.Enabled = False
    End If
    If frmLogin.Check2(6).Value = 1 Then
        Form1.Impresion.Enabled = True
    Else
        Form1.Impresion.Enabled = False
    End If
    If frmLogin.Check2(7).Value = 1 Then
        Form1.Audiometria.Enabled = True
    Else
        Form1.Audiometria.Enabled = False
    End If
    If frmLogin.Check2(8).Value = 1 Then
        Form1.Tuberculosis.Enabled = True
    Else
        Form1.Tuberculosis.Enabled = False
    End If
    If frmLogin.Check2(9).Value = 1 Then
        Form1.Cardiologia.Enabled = True
    Else
        Form1.Cardiologia.Enabled = False
    End If
    If frmLogin.Check3(0).Value = 1 Then
        Form1.Editar_eliminar_registros.Enabled = True
    Else
        Form1.Editar_eliminar_registros.Enabled = False
    End If
    If frmLogin.Check3(1).Value = 1 Then
        Form1.Importar_Somatometria.Enabled = True
    Else
        Form1.Importar_Somatometria.Enabled = False
    End If
    If frmLogin.Check3(2).Value = 1 Then
        Form1.Exportar_informacion.Enabled = True
    Else
        Form1.Exportar_informacion.Enabled = False
    End If
    If frmLogin.Check4(0).Value = 1 Then
        Form1.Familiares.Enabled = True
    Else
        Form1.Familiares.Enabled = False
    End If
    If frmLogin.Check4(1).Value = 1 Then
        Form1.Estatisticas.Enabled = True
    Else
        Form1.Estatisticas.Enabled = False
    End If
End Sub
Private Sub txtUserName_Change()
    On Error Resume Next
    With RRsUsers
        .Requery
        If OPTION1.Value = True Then
            .Filter = "NOMBRE LIKE '*" & txtUserName & "*'"
            If frmLogin.Check1.Value = 1 Then
                Form1.Usuario.Enabled = True
            Else
                Form1.Usuario.Enabled = False
            End If
            If frmLogin.Check2(0).Value = 1 Then
                Form1.Somatometria.Enabled = True
            Else
                Form1.Somatometria.Enabled = False
            End If
            If frmLogin.Check2(1).Value = 1 Then
                Form1.Laboratorio.Enabled = True
            Else
                Form1.Laboratorio.Enabled = False
            End If
            If frmLogin.Check2(2).Value = 1 Then
                Form1.Dental.Enabled = True
            Else
                Form1.Dental.Enabled = False
            End If
            If frmLogin.Check2(3).Value = 1 Then
                Form1.Nutricion.Enabled = True
            Else
                Form1.Nutricion.Enabled = False
            End If
            If frmLogin.Check2(4).Value = 1 Then
                Form1.Salud_mujer.Enabled = True
            Else
                Form1.Salud_mujer.Enabled = False
            End If
            If frmLogin.Check2(5).Value = 1 Then
                Form1.Optometría.Enabled = True
            Else
                Form1.Optometría.Enabled = False
            End If
            If frmLogin.Check2(6).Value = 1 Then
                Form1.Impresion.Enabled = True
            Else
                Form1.Impresion.Enabled = False
            End If
            If frmLogin.Check2(7).Value = 1 Then
                Form1.Audiometria.Enabled = True
            Else
                Form1.Audiometria.Enabled = False
            End If
            If frmLogin.Check2(8).Value = 1 Then
                Form1.Tuberculosis.Enabled = True
            Else
                Form1.Tuberculosis.Enabled = False
            End If
            If frmLogin.Check2(9).Value = 1 Then
                Form1.Cardiologia.Enabled = True
            Else
                Form1.Cardiologia.Enabled = False
            End If
            If frmLogin.Check3(0).Value = 1 Then
                Form1.Editar_eliminar_registros.Enabled = True
            Else
                Form1.Editar_eliminar_registros.Enabled = False
            End If
            If frmLogin.Check3(1).Value = 1 Then
                Form1.Importar_Somatometria.Enabled = True
            Else
                Form1.Importar_Somatometria.Enabled = False
            End If
            If frmLogin.Check3(2).Value = 1 Then
                Form1.Exportar_informacion.Enabled = True
            Else
                Form1.Exportar_informacion.Enabled = False
            End If
            If frmLogin.Check4(0).Value = 1 Then
                Form1.Familiares.Enabled = True
            Else
                Form1.Familiares.Enabled = False
            End If
            If frmLogin.Check4(1).Value = 1 Then
                Form1.Estatisticas.Enabled = True
            Else
                Form1.Estatisticas.Enabled = False
            End If
        Else
            .Filter = ""
            .MoveFirst
        End If
    End With
End Sub
