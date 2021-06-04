VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Punto de Venta"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Salir"
      Height          =   500
      Index           =   7
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Control de saldos"
      Height          =   500
      Index           =   6
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Control de ventas"
      Height          =   500
      Index           =   5
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Ventas"
      Height          =   500
      Index           =   4
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Control de pagos"
      Height          =   500
      Index           =   3
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Pagos"
      Height          =   500
      Index           =   2
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Clientes"
      Height          =   500
      Index           =   1
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009EC0C2&
      Caption         =   "Articulos"
      Height          =   500
      Index           =   0
      Left            =   720
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   5595
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   7440
      TabIndex        =   10
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   6375
         Begin VB.CommandButton Command2 
            BackColor       =   &H009EC0C2&
            Caption         =   "P"
            Height          =   495
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   5160
            Width           =   495
         End
         Begin VB.Image Image1 
            Height          =   5535
            Left            =   360
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5655
         End
      End
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
            TipoCatalogo = 0
            TipoConsulta = 0
            Form2.Show
            Form2.Text1.SetFocus
            Form1.Enabled = False
        Case 1
            TipoCatalogo = 1
            TipoConsulta = 0
            Form2.Show
            Form2.Text1.SetFocus
            Form1.Enabled = False
        Case 2
            TipoCatalogo = 1
            TipoConsulta = 1
            TipoTransaccion = 0
            Form3.Show
            Form3.DTPicker1.SetFocus
            Form1.Enabled = False
        Case 3
            TipoCatalogo = 2
            Form5.Show
            Form5.DTPicker1(0).SetFocus
            Form1.Enabled = False
        Case 4
            TipoConsulta = 1
            TipoTransaccion = 1
            Form4.Show
            Form4.DTPicker1.SetFocus
            Form1.Enabled = False
        Case 5
            TipoCatalogo = 3
            Form5.Show
            Form5.DTPicker1(0).SetFocus
            Form1.Enabled = False
        Case 6
            Form6.Show
            Form1.Enabled = False
        Case 7
            RsIdVenta.Close
            RsClientes.Close
            RsItems.Close
            RsPagos.Close
            RsVentas.Close
            RsPagosV.Close
            RsSaldosV.Close
            RsVentasV.Close
            RsVentasTxt.Close
            RsCabeceraVentas.Close
            RsIdPagos.Close
            Cn.Close
            Unload Me
    End Select

End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case 0
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(0).Caption & ".jpg")
        Case 1
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(1).Caption & ".jpg")
        Case 2
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(2).Caption & ".jpg")
        Case 3
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(3).Caption & ".jpg")
        Case 4
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(4).Caption & ".jpg")
        Case 5
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(5).Caption & ".jpg")
        Case 6
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(6).Caption & ".jpg")
        Case 7
            Image1.Picture = LoadPicture(App.Path & "\Images\" & Command1(7).Caption & ".jpg")
    End Select

End Sub

Private Sub Command2_Click()

    Form9.Show
    Form1.Enabled = False

End Sub

Private Sub Form_Load()

    If PrPreferencias = 0 Then
        Command2.Visible = False
    Else
        Command2.Visible = True
    End If
    Image1.Picture = LoadPicture(App.Path & "\Images\Inicio.jpg")
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image1.Picture = LoadPicture(App.Path & "\Images\Inicio.jpg")

End Sub
