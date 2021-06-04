VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "JAHG Software - Control Bovino"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5640
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
   ScaleHeight     =   4680
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   4200
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu SeleccionarProductor 
         Caption         =   "Seleccionar Productor"
      End
      Begin VB.Menu SeleccionarHato 
         Caption         =   "Seleccionar Hato"
         Enabled         =   0   'False
      End
      Begin VB.Menu SeleccionarImagenDeFondo 
         Caption         =   "Seleccionar Imagen de Fondo"
      End
      Begin VB.Menu RespaldarBaseDeDatos 
         Caption         =   "Respaldar Base de Datos"
      End
      Begin VB.Menu RestaurarBaseDeDatos 
         Caption         =   "Restaurar Base de Datos"
      End
      Begin VB.Menu SalirDelPrograma 
         Caption         =   "Salir del Programa"
      End
   End
   Begin VB.Menu General 
      Caption         =   "General"
      Begin VB.Menu Productor 
         Caption         =   "Productor"
      End
      Begin VB.Menu Hato 
         Caption         =   "Hato"
         Enabled         =   0   'False
      End
      Begin VB.Menu Personal 
         Caption         =   "Personal"
         Enabled         =   0   'False
      End
      Begin VB.Menu ListasDeValores 
         Caption         =   "Listas de Valores"
      End
   End
   Begin VB.Menu Animales 
      Caption         =   "Animales"
      Begin VB.Menu CargarEventos 
         Caption         =   "Cargar Eventos"
      End
      Begin VB.Menu GeneralAnimales 
         Caption         =   "General Animales"
      End
      Begin VB.Menu Fichas 
         Caption         =   "Fichas"
      End
      Begin VB.Menu Loteo 
         Caption         =   "Loteo"
      End
      Begin VB.Menu Toros 
         Caption         =   "Toros"
         Begin VB.Menu StockLoteo 
            Caption         =   "Stock Loteo"
         End
         Begin VB.Menu ConsultarPadron 
            Caption         =   "Consultar Padròn"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    ProcFormResize
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ProcSalir
End Sub
Private Sub GeneralAnimales_Click()
    ProcConexionAnimales
End Sub
Private Sub Hato_Click()
    ProcHatos
End Sub
Private Sub Label2_Change()
    ProcHabilitarHato
End Sub
Private Sub ListasDeValores_Click()
    VarForm5State = 1
    Form5.Show
End Sub
Private Sub Personal_Click()
    ProcPersonales
End Sub
Private Sub Productor_Click()
    ProcProductores
End Sub
Private Sub RespaldarBaseDeDatos_Click()
    ProcRespaldar
End Sub
Private Sub RestaurarBaseDeDatos_Click()
    ProcRestaurar
End Sub
Private Sub SalirDelPrograma_Click()
    ProcSalir
End Sub
Private Sub SeleccionarHato_Click()
    ProcSeleccionHato
End Sub
Private Sub SeleccionarImagenDeFondo_Click()
    ProcSeleccionarImagenFondo
End Sub
Private Sub SeleccionarProductor_Click()
    ProcSeleccionProductor
End Sub
Private Sub Timer1_Timer()
    ProcReloj
End Sub
