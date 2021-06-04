VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Semana de la salud"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   11925
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   10335
      Left            =   -120
      Picture         =   "Form1.frx":324A
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   12135
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Usuario 
         Caption         =   "Usuarios"
         Shortcut        =   ^U
      End
      Begin VB.Menu CerrarSesion 
         Caption         =   "Cerrar sesión"
         Shortcut        =   ^W
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu Modulos 
      Caption         =   "Módulos"
      Begin VB.Menu Somatometria 
         Caption         =   "Somatometría"
         Begin VB.Menu Nuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^V
         End
         Begin VB.Menu Registrado 
            Caption         =   "Registrado"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu Laboratorio 
         Caption         =   "Laboratorio"
         Shortcut        =   ^L
      End
      Begin VB.Menu Dental 
         Caption         =   "Dental"
         Shortcut        =   ^D
      End
      Begin VB.Menu Nutricion 
         Caption         =   "Nutrición"
         Shortcut        =   ^N
      End
      Begin VB.Menu Salud_mujer 
         Caption         =   "Salud de la mujer"
         Shortcut        =   ^M
      End
      Begin VB.Menu Optometría 
         Caption         =   "Optometría"
         Shortcut        =   ^O
      End
      Begin VB.Menu Audiometria 
         Caption         =   "Audiometría"
         Shortcut        =   ^A
      End
      Begin VB.Menu Tuberculosis 
         Caption         =   "Tuberculosis"
         Shortcut        =   ^T
      End
      Begin VB.Menu Cardiologia 
         Caption         =   "Cardiologia"
         Shortcut        =   ^C
      End
      Begin VB.Menu Impresion 
         Caption         =   "Impresión de  resultados"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu Edicion 
      Caption         =   "Edición"
      Begin VB.Menu Editar_eliminar_registros 
         Caption         =   "Eliminar o actualizar información"
         Shortcut        =   ^E
      End
      Begin VB.Menu Datos_externos 
         Caption         =   "Datos externos"
         Begin VB.Menu Importar_Somatometria 
            Caption         =   "Importar Somatometria (Archivo .xls)"
            Shortcut        =   ^I
         End
         Begin VB.Menu Exportar_informacion 
            Caption         =   "Exportar toda la información (Archivo .xls)"
            Shortcut        =   ^X
         End
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu Familiares 
         Caption         =   "Familiares"
         Shortcut        =   ^F
      End
      Begin VB.Menu Estatisticas 
         Caption         =   "Estadisticas"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu Acerca 
         Caption         =   "Acerca"
         Shortcut        =   +{INSERT}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acerca_Click()
    On Error Resume Next
    frmAbout.Show
    Me.Enabled = False
End Sub

Private Sub Audiometria_Click()
    On Error Resume Next
    FormAudiometria.Show
    Me.Enabled = False
End Sub

Private Sub Cardiologia_Click()
    On Error Resume Next
    FormCardiologia.Show
    Me.Enabled = False
End Sub
Private Sub CerrarSesion_Click()
    On Error Resume Next
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
    Form1.Optometría.Enabled = False
    Form1.Tuberculosis.Enabled = False
    Form1.Cardiologia.Enabled = False
    frmLogin.Show
    frmLogin.txtUserName.Text = ""
    frmLogin.txtPassword.Text = ""
    frmLogin.txtUserName.SetFocus
    Me.Enabled = False
End Sub
Private Sub Dental_Click()
    On Error Resume Next
    FormDental.Show
    Me.Enabled = False
End Sub
Private Sub Editar_eliminar_registros_Click()
    On Error Resume Next
    FormEditar.Show
    Me.Enabled = False
End Sub
Private Sub Estatisticas_Click()
    On Error Resume Next
    FEstadisticas.Show
    Me.Enabled = False
End Sub
Private Sub Exportar_informacion_Click()
    On Error Resume Next
    FormExportarTodo.Show
    Me.Enabled = False
End Sub
Private Sub Familiares_Click()
    On Error Resume Next
    ReporteFamiliares.Show
    Me.Enabled = False
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RsUsers
        If .State = 1 Then .Close
            .Open "Select * from USERS", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
End Sub
Private Sub Importar_Somatometria_Click()
    On Error Resume Next
    FormImportarSomatometria.Show
    Me.Enabled = False
End Sub
Private Sub Impresion_Click()
    On Error Resume Next
    FormImpresion.Show
    Me.Enabled = False
End Sub
Private Sub Laboratorio_Click()
    On Error Resume Next
    FormLaboratorio.Show
    Me.Enabled = False
End Sub
Private Sub Nuevo_Click()
    On Error Resume Next
    FormSomatometriaNuevo.Show
    Me.Enabled = False
End Sub
Private Sub Nutricion_Click()
    On Error Resume Next
    FormNutricion.Show
    Me.Enabled = False
End Sub
Private Sub Optometría_Click()
    On Error Resume Next
    FormOptometria.Show
    Me.Enabled = False
End Sub
Private Sub Registrado_Click()
    On Error Resume Next
    FormSomatometriaRegistrado.Show
    Me.Enabled = False
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    Unload FormSomatometriaRegistrado
    Unload FormSomatometriaNuevo
    Unload BuscarTrabajador
    Unload FormLaboratorio
    Unload FormDental
    Unload FormNutricion
    Unload FormMujer
    Unload FormOptometria
    Unload FormImpresion
    Unload DataReport1
    Unload FormImportarSomatometria
    Unload FormEditar
    Unload FormExportarTodo
    Unload DataReport2
    Unload ReporteFamiliares
    Unload Usuarios
    Unload frmLogin
    Unload frmAbout
    Unload FEstadisticas
    Unload FormAudiometria
    Unload FormTuberculosis
    Unload FormCardiologia
    Unload Form1
End Sub
Private Sub Salud_mujer_Click()
    On Error Resume Next
    FormMujer.Show
    Me.Enabled = False
End Sub

Private Sub Tuberculosis_Click()
    On Error Resume Next
    FormTuberculosis.Show
    Me.Enabled = False
End Sub
Private Sub Usuario_Click()
    On Error Resume Next
    Me.Enabled = False
    Usuarios.Show
    Set Usuarios.Text1(2).DataSource = RsUsers
    Set Usuarios.Text1(3).DataSource = RsUsers
    Set Usuarios.Check7(0).DataSource = RsUsers
    Set Usuarios.Check8(0).DataSource = RsUsers
    Set Usuarios.Check8(1).DataSource = RsUsers
    Set Usuarios.Check8(2).DataSource = RsUsers
    Set Usuarios.Check8(3).DataSource = RsUsers
    Set Usuarios.Check8(4).DataSource = RsUsers
    Set Usuarios.Check8(5).DataSource = RsUsers
    Set Usuarios.Check8(6).DataSource = RsUsers
    Set Usuarios.Check9(0).DataSource = RsUsers
    Set Usuarios.Check9(1).DataSource = RsUsers
    Set Usuarios.Check9(2).DataSource = RsUsers
    Set Usuarios.Check10(0).DataSource = RsUsers
    Set Usuarios.Check10(1).DataSource = RsUsers
    Usuarios.Text1(2).DataField = ("NOMBRE")
    Usuarios.Text1(3).DataField = ("PASS")
    Usuarios.Check7(0).DataField = ("USUARIOS")
    Usuarios.Check8(0).DataField = ("SOMATOMETRIA")
    Usuarios.Check8(1).DataField = ("LABORATORIO")
    Usuarios.Check8(2).DataField = ("DENTAL")
    Usuarios.Check8(3).DataField = ("NUTRICION")
    Usuarios.Check8(4).DataField = ("MUJER")
    Usuarios.Check8(5).DataField = ("OPTOMETRIA")
    Usuarios.Check8(6).DataField = ("IMPRESION")
    Usuarios.Check9(0).DataField = ("EDITAR")
    Usuarios.Check9(1).DataField = ("IMPORTAR")
    Usuarios.Check9(2).DataField = ("EXPORTAR")
    Usuarios.Check10(0).DataField = ("FAMILIARES")
    Usuarios.Check10(0).DataField = ("ESTADISTICAS")
End Sub
