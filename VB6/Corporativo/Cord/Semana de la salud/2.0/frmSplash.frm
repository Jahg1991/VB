VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   40
      Left            =   1320
      Top             =   4200
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   4260
      TabIndex        =   0
      Top             =   5640
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   7080
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   0
      Picture         =   "frmSplash.frx":324A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    On Error Resume Next
    Form1.Show
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
    Form1.Audiometria.Enabled = False
    Form1.Tuberculosis.Enabled = False
    Form1.Cardiologia.Enabled = False
    Form1.Enabled = False
    frmLogin.Show
    Unload frmSplash
End Sub
Private Sub Timer2_Timer()
    On Error Resume Next
    If pb.Value = 100 Then
        pb.Value = 0
    Else
        pb.Value = Val(pb.Value) + 1
    End If
End Sub
