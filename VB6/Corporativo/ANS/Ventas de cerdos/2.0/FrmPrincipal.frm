VERSION 5.00
Begin VB.Form FrmPrincipal 
   BorderStyle     =   0  'None
   Caption         =   "Venta de cerdos"
   ClientHeight    =   11295
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   17925
   ControlBox      =   0   'False
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11295
   ScaleWidth      =   17925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   11295
      Left            =   0
      Picture         =   "FrmPrincipal.frx":324A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17895
   End
   Begin VB.Menu Ingreso 
      Caption         =   "Ingreso"
      Begin VB.Menu Nuevo_usuario 
         Caption         =   "Agregar, editar o eliminar usuarios"
      End
      Begin VB.Menu cerrar_sesion 
         Caption         =   "Cerrar sesiòn"
      End
   End
   Begin VB.Menu Captura 
      Caption         =   "Captura"
      Begin VB.Menu nueva_venta 
         Caption         =   "Nueva venta"
      End
      Begin VB.Menu editar_eliminar_venta 
         Caption         =   "Editar o eliminar venta"
      End
   End
   Begin VB.Menu Consulta 
      Caption         =   "Consulta"
      Begin VB.Menu Cliente 
         Caption         =   "Cliente"
      End
      Begin VB.Menu Granja 
         Caption         =   "Granja"
      End
      Begin VB.Menu Fecha 
         Caption         =   "Fecha"
      End
   End
   Begin VB.Menu Acerca 
      Caption         =   "Acerca"
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acerca_Click()
On Error Resume Next
frmAbout.Show
End Sub

Private Sub cerrar_sesion_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Cliente_Click()
On Error Resume Next
frmcon_nombre.Show
End Sub

Private Sub editar_eliminar_venta_Click()
On Error Resume Next
frmMod_elim_ventas.Show
End Sub

Private Sub Fecha_Click()
On Error Resume Next
frmcon_fecha.Show
End Sub

Private Sub Form_Resize()
On Error Resume Next
Image1.Height = FrmPrincipal.Height
Image1.Width = FrmPrincipal.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim i As Integer
While Forms.Count > 1
i = 0
While Forms(i).Caption = Me.Caption
i = i + 1
Wend
Unload Forms(i)
Wend
Unload Me
End
End Sub

Private Sub Granja_Click()
On Error Resume Next
frmcon_granja.Show
End Sub

Private Sub nueva_venta_Click()
On Error Resume Next
frmCaptura.Show
End Sub

Private Sub Nuevo_usuario_Click()
On Error Resume Next
Load frmNuevo_usuario
frmNuevo_usuario.Show
End Sub

Private Sub Salir_Click()
On Error Resume Next
Unload Me
End Sub
