VERSION 5.00
Begin VB.Form frmCompras 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Nueva Compra"
   ClientHeight    =   9075
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   17415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   17175
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8655
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   16935
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ELIMINAR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   3960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "AÑADIR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   10
            Left            =   2280
            TabIndex        =   8
            Top             =   2600
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Index           =   0
            Left            =   2280
            TabIndex        =   2
            Top             =   600
            Width           =   14535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   2280
            TabIndex        =   1
            Top             =   120
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Index           =   1
            Left            =   2280
            TabIndex        =   3
            Top             =   1120
            Width           =   14535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   1
            Left            =   2280
            MaxLength       =   7
            TabIndex        =   4
            Top             =   1640
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   2
            Left            =   2280
            MaxLength       =   7
            TabIndex        =   5
            Top             =   2120
            Width           =   3015
         End
         Begin VB.ListBox List1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   2055
            Left            =   120
            TabIndex        =   7
            Top             =   4920
            Width           =   16695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   3
            Left            =   2280
            TabIndex        =   6
            Top             =   3080
            Width           =   14535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   5
            Left            =   10320
            TabIndex        =   10
            Top             =   7680
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   6
            Left            =   10320
            TabIndex        =   9
            Top             =   7200
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   4
            Left            =   10320
            TabIndex        =   11
            Top             =   8160
            Width           =   6495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LOTE"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   6
            Left            =   -480
            TabIndex        =   31
            Top             =   2600
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "FOLIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   14
            Left            =   1080
            TabIndex        =   30
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NOMBRE"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   29
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ARTICULO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   28
            Top             =   1120
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   27
            Top             =   1640
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   5
            Left            =   -480
            TabIndex        =   26
            Top             =   2120
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   9
            Left            =   8040
            TabIndex        =   25
            Top             =   8160
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COMENTARIOS"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   24
            Top             =   3080
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   7
            Left            =   8040
            TabIndex        =   23
            Top             =   7680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SUBTOTAL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   10
            Left            =   8040
            TabIndex        =   22
            Top             =   7200
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmCompras.frx":0000
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   4560
            Width           =   16575
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   3240
      Index           =   0
      Left            =   5100
      TabIndex        =   0
      Top             =   2917
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3015
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   6975
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ACEPTAR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   0
            Left            =   1560
            TabIndex        =   12
            Top             =   120
            Width           =   5055
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   1
            Left            =   1560
            TabIndex        =   13
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   2
            Left            =   1560
            TabIndex        =   14
            Top             =   1560
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   11
            Left            =   -720
            TabIndex        =   18
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PAGADO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   12
            Left            =   -720
            TabIndex        =   17
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CAMBIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   13
            Left            =   -720
            TabIndex        =   16
            Top             =   1680
            Width           =   2055
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Anadir 
         Caption         =   "Añadir Proveedor"
         Shortcut        =   ^A
      End
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmCompras
'Proposito:     Registro de Compras
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        25/05/2021     Alfredo Hernandez    Se agrego usuario fecha de
'                                               modificacion y fecha de creacion a
'                                               todos los insert
'
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset  'folio
    Dim RS1                 As New adodb.Recordset  'clientesproveedores
    Dim Rs2                 As New adodb.Recordset  'items
    Dim Rs3                 As New adodb.Recordset  'ventascompras
    Dim Rs4                 As New adodb.Recordset  'lista de ingredientes
    Dim Rs5                 As New adodb.Recordset  'movimientos de inventarios
    Dim Rs6                 As New adodb.Recordset  'ticket
    Dim Rs7                 As New adodb.Recordset  'lotes
    Dim Rs9                 As New adodb.Recordset  'movimientos de caja
    '//OTROS
    Dim TipoErr             As Long
    Dim i                   As Long
    Dim X                   As Long
    Dim intX                As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim c3                  As Long
    Dim c4                  As Long
    Dim Prt                 As Printer
    Dim vCategoria          As String
    '//VALORES PARA INSERTAR
    Dim v1                  As Long                 'idclienteproveedor
    Dim v2                  As String               'nombre cliente proveedor
    Dim v3                  As String               'folio
    Dim v4                  As String               'LugarVenta
    Dim v5                  As String               'mesa
    Dim v6                  As Date                 'fecha
    Dim v7                  As String               'Tipoarticulo
    Dim v8                  As Long                 'idarticulo
    Dim v9                  As String               'codigo articulo
    Dim v10                 As String               'descripcion articulo
    Dim v11                 As String               'cantidad
    Dim v12                 As String               'UDM
    Dim v13                 As String               'precio
    Dim v14                 As String               'subtotal
    Dim v15                 As String               'iva
    Dim v16                 As String               'total
    Dim v17                 As String               'totalpagado
    Dim v18                 As String               'cancelado
    Dim v19                 As String               'comentarios
    Dim v20                 As String               'tipo
    Dim v21                 As String               'lote
    Dim IdTransaccion       As Long                 'folio
    '//LOTES
    Dim ControlLote         As Boolean
    Dim InItemId            As Long
    Dim InLoteExiste        As Long
    '//COMPRAS
    Dim listSubtotal        As String
    Dim listIva             As String
    Dim listTotal           As String
    Dim vLstCantidad        As String
    Dim vLstPrecio          As String
    Dim viva                As String
    Dim viid                As String
    Dim videscripcion       As String
    Dim vicantidad          As String
    Dim viprecio            As String
    Dim DineroRestante      As String
    Dim vTicketSubtotal     As String
    Dim vTicketIva          As String
    Dim vTicketTotal        As String
    '//CREDITO
    Dim Credito             As String               'credito del cliente
    Dim CreditoUsado        As String               'Credito usado del cliente
    Dim DiasCredito         As Long                 'Dias de credito del cliente
    Dim DiasCreditoUsado    As Long                 'Dias de credito usados por el cliente
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        If StPermisosCatalogos = "Si" Then
            With Anadir
                .Visible = True
            End With
        Else
            With Anadir
                .Visible = False
            End With
        End If
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from PO_TRANSACTION_ID_P", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            .MoveFirst
            If IsNull(Rs!IdCompra) = False Then
                IdTransaccion = Rs!IdCompra
            Else
                IdTransaccion = 1
            End If
        End With
        
        With Text1(0)
            .Text = "C-" & IdTransaccion
        End With
        
        For i = 1 To 3
            With Text1(i)
                .BackColor = COLOR_NO_ENCONTRADO
            End With
        Next i
        
        For i = 0 To 1
            With Combo1(i)
                .BackColor = COLOR_NO_ENCONTRADO
            End With
        Next i
        
        With RS1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from HZ_PARTY where proveedor = 'Si' order by 2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            If .RecordCount <> 0 Then
                With Combo1(0)
                    .Clear
                End With
                
                While Not .EOF
                    Combo1(0).AddItem .Fields(1)
                    .MoveNext
                Wend
            End If
        End With
        With Rs2
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_SYSTEM_ITEMS order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            If .RecordCount <> 0 Then
                With Combo1(1)
                    .Clear
                End With
                
                While Not .EOF
                    Combo1(1).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                    .MoveNext
                Wend
            End If
        End With
        
        With Rs2
            If .RecordCount = 0 Then
                MsgBox "No hay registros existentes", vbOKOnly, "Información"
                Exit Sub
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_Click(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                With Combo1(0)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                        With RS1
                            .Filter = ""
                            .Requery
                            v1 = 0
                            v2 = ""
                        End With
                    Else
                        .BackColor = COLOR_NORMAL
                        With RS1
                            .Filter = ""
                            .Filter = "nombre like '" & Combo1(0) & "'"
                            .Requery
                            v1 = .Fields(0)
                            v2 = .Fields(1)
                        End With
                    End If
                End With
            Case 1
                With Combo1(1)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                        With Rs2
                            .Filter = ""
                            .Requery
                        End With
                        
                        With Text1(2)
                            .Text = ""
                        End With
                        
                        With Text1(10)
                            .Text = ""
                        End With
                    Else
                        .BackColor = COLOR_NORMAL
                        InItemId = Get_ItemId(.Text)
                        With Rs2
                            .Filter = "Id = " & InItemId
                            .Requery
                        End With
                        
                        With Rs2
                            Text1(2).Text = Replace(.Fields(7).Value, ",", ".")
                            If .Fields(11).Value = 1 Then
                                With Text1(10)
                                    .Text = "C" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdTransaccion
                                End With
                            Else
                                With Text1(10)
                                    .Text = ""
                                End With
                            End If
                        End With
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        On Error GoTo errHandler
        Static cadena As String
        Select Case Index
            Case 0
                With Combo1(0)
                    ' si pesionamos las teclas de las flechas sale de la rutina
                    If KeyCode = vbKeyUp Then Exit Sub
                    
                    If KeyCode = vbKeyDown Then Exit Sub
                    
                    If KeyCode = vbKeyLeft Then Exit Sub
                    
                    If KeyCode = vbKeyRight Then Exit Sub
                    
                    ' verifica qu no se presionó la tecla backspace
                    If KeyCode <> vbKeyBack Then
                        cadena = Mid(.Text, 1, Len(.Text) - .SelLength)
                    Else
                        '...tecla backspace
                        If cadena <> "" Then
                            cadena = Mid(cadena, 1, Len(cadena) - 1)
                        End If
                    End If
                    
                    For i = 0 To .ListCount - 1
                        If UCase(cadena) = UCase(Mid(.List(i), 1, Len(cadena))) Then
                            .ListIndex = i
                            Exit For
                        End If
                    Next
                    ' Seelecciona
                    .SelStart = Len(cadena)
                    .SelLength = Len(.Text)
                    If .ListIndex = -1 Then
                        ' color de fondo del combo en caso de que no hay coincidencias
                        .BackColor = COLOR_NO_ENCONTRADO
                        With RS1
                            .Filter = ""
                            .Requery
                            v1 = 0
                            v2 = ""
                        End With
                    Else
                        ' Backcolor normal cuando hay coincidencia
                        .BackColor = COLOR_NORMAL
                        With RS1
                            .Filter = ""
                            .Filter = "nombre like '" & Combo1(0) & "'"
                            .Requery
                            v1 = .Fields(0)
                            v2 = .Fields(1)
                        End With
                    End If
                End With
            Case 1
                With Combo1(1)
                    ' si pesionamos las teclas de las flechas sale de la rutina
                    If KeyCode = vbKeyUp Then Exit Sub
                    
                    If KeyCode = vbKeyDown Then Exit Sub
                    
                    If KeyCode = vbKeyLeft Then Exit Sub
                    
                    If KeyCode = vbKeyRight Then Exit Sub
                    
                    ' verifica qu no se presionó la tecla backspace
                    If KeyCode <> vbKeyBack Then
                        cadena = Mid(.Text, 1, Len(.Text) - .SelLength)
                    Else
                        '...tecla backspace
                        If cadena <> "" Then
                            cadena = Mid(cadena, 1, Len(cadena) - 1)
                        End If
                    End If
                    
                    For i = 0 To .ListCount - 1
                        If UCase(cadena) = UCase(Mid(.List(i), 1, Len(cadena))) Then
                            .ListIndex = i
                            Exit For
                        End If
                    Next
                    ' Seelecciona
                    .SelStart = Len(cadena)
                    .SelLength = Len(.Text)
                    If .ListIndex = -1 Then
                        ' color de fondo del combo en caso de que no hay coincidencias
                        .BackColor = COLOR_NO_ENCONTRADO
                        With Rs2
                            .Filter = ""
                            .Requery
                        End With
                        
                        With Text1(2)
                            .Text = ""
                        End With
                        
                        With Text1(10)
                            .Text = ""
                        End With
                    Else
                        ' Backcolor normal cuando hay coincidencia
                        .BackColor = COLOR_NORMAL
                        InItemId = Get_ItemId(.Text)
                        With Rs2
                            .Filter = "Id = " & InItemId
                            .Requery
                            Text1(2).Text = Replace(.Fields(7).Value, ",", ".")
                            If .Fields(11).Value = 1 Then
                                With Text1(10)
                                    .Text = "C" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdTransaccion
                                End With
                            Else
                                With Text1(10)
                                    .Text = ""
                                End With
                            End If
                        End With
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                If Combo1(1) <> "" And Text1(1) <> "" And Text1(2) <> "" Then
                    With Rs2
                        viid = .Fields(0).Value
                    End With
                    
                    With Combo1(1)
                        videscripcion = Mid(.Text, 1, 47)
                    End With
                    
                    With Text1(1)
                        vicantidad = Replace(Format(Val(.Text), "0.00"), ",", ".")
                    End With
                    
                    With Text1(2)
                        viprecio = Replace(Format(Val(.Text), "0.00"), ",", ".")
                    End With
                    ' 1 - 10
                    c1 = 10 - Len(viid)
                    For i = 1 To c1
                        viid = " " & viid
                    Next i
                    ' 12 - 58
                    c2 = 47 - Len(videscripcion)
                    For i = 1 To c2
                        videscripcion = videscripcion & " "
                    Next i
                    ' 60 - 74
                    c3 = 15 - Len(vicantidad)
                    For i = 1 To c3
                        vicantidad = " " & vicantidad
                    Next i
                    ' 76 - 90
                    c4 = 15 - Len(viprecio)
                    For i = 1 To c4
                        viprecio = " " & viprecio
                    Next i
                    
                    With List1
                        .AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                    End With
                    
                    With Text1(1)
                        .Text = ""
                    End With
                    
                    With Text1(2)
                        .Text = ""
                    End With
                    
                    With Combo1(1)
                        .Text = ""
                        .BackColor = COLOR_NO_ENCONTRADO
                        .SetFocus
                    End With
                    listSubtotal = 0
                    listIva = 0
                    With List1
                        For i = 0 To .ListCount - 1
                            .ListIndex = i
                            .SetFocus
                            vLstCantidad = Trim(Mid(.Text, 60, 15))
                            vLstPrecio = Trim(Mid(.Text, 76, 15))
                            viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                            listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                            listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                        Next i
                    End With
                    listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                    listIva = Replace(Format(listIva, "0.00"), ",", ".")
                    listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                    listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                    With Text1(6)
                        .Text = listSubtotal
                    End With
                    
                    With Text1(5)
                        .Text = listIva
                    End With
                    
                    With Text1(4)
                        .Text = listTotal
                    End With
                Else
                    MsgBox "Llenar todos los campos", vbCritical, "Error"
                    With Combo1(1)
                        .SetFocus
                    End With
                End If
            Case 1
                With List1
                    intX = .ListIndex
                    .RemoveItem intX
                    listSubtotal = 0
                    listIva = 0
                    For i = 0 To .ListCount - 1
                        .ListIndex = i
                        .SetFocus
                        vLstCantidad = Trim(Mid(.Text, 60, 15))
                        vLstPrecio = Trim(Mid(.Text, 76, 15))
                        viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                        listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                        listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                    Next i
                End With
                listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                listIva = Replace(Format(listIva, "0.00"), ",", ".")
                listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                With Text1(6)
                    .Text = listSubtotal
                End With
                
                With Text1(5)
                    .Text = listIva
                End With
                
                With Text1(4)
                    .Text = listTotal
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Text1_Change(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 1
                With Text1(1)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 2
                With Text1(2)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 3
                With Text1(3)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Text2_Change(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                With Text2(2)
                    .Text = Replace(Format(Val(Text2(1)) - Val(Text2(0)), "0.00"), ",", ".")
                End With
            Case 1
                With Text2(2)
                    .Text = Replace(Format(Val(Text2(1)) - Val(Text2(0)), "0.00"), ",", ".")
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Text2_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Anadir_Click()
        On Error GoTo errHandler
        InTipoAltaClienteProveedor = 1
        StTipoClienteProveedor = "Proveedor"
        With frmProveedoresNuevo
            .Caption = "Añadir nuevo Proveedor"
            .Show 1
        End With
        Unload frmCompras
        Set frmCompras = Nothing
        
        With frmCompras
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Anadir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        With List1
            If .ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
                'Mostrar pago
                With Frame1(0)
                    .Enabled = False
                End With
                
                With Frame2(0)
                    .Visible = True
                End With
                
                With Archivo
                    .Enabled = False
                End With
                
                With Text2(0)
                    .Text = Text1(4)
                End With
            Else
                MsgBox "Llenar todos los campos", vbCritical, "Advertencia"
                Exit Sub
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
        If vbq = vbYes Then
            Credito = Get_Credito(v1)
            CreditoUsado = Get_CreditoUsado(v1)
            DiasCredito = Get_CreditoDias(v1)
            DiasCreditoUsado = Get_CreditoDiasUsado(v1)
            If Val(Text2(1)) = 0 Or Text2(1) = "" Then
                With Text2(1)
                    .Text = "0"
                End With
            End If
            'credito
            If Val(CreditoUsado) + Val(Text2(0)) - Val(Text2(1)) > Val(Credito) Then
                MsgBox "No se puede realizar la compra, ya supero su limite de crédito", vbCritical, "Error"
                'Ocultar pago
                With Frame1(0)
                    .Enabled = True
                End With
                
                With Frame2(0)
                    .Visible = False
                End With
                
                With Archivo
                    .Enabled = True
                End With
                Exit Sub
            End If
            If DiasCreditoUsado > DiasCredito And Val(Text2(0)) - Val(Text2(1)) <> 0 Then
                MsgBox "No se puede realizar la compra, ya supero sus días de crédito", vbCritical, "Error"
                'Ocultar pago
                With Frame1(0)
                    .Enabled = True
                End With
                
                With Frame2(0)
                    .Visible = False
                End With
                
                With Archivo
                    .Enabled = True
                End With
                
                Exit Sub
            End If
            'Actualizar Folio
            With Rs
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from PO_TRANSACTION_ID_P", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                .MoveFirst
                If IsNull(Rs!IdCompra) = False Then
                    IdTransaccion = Rs!IdCompra
                Else
                    IdTransaccion = 1
                End If
                
                With Text1(0)
                    .Text = "C-" & IdTransaccion
                End With
            End With
            
            With Text2(1)
                If Val(.Text) > 0 Then
                    With Rs9
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from RA_CASH_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                            With .Fields(1)
                                .Value = Date                                                           'fecha
                            End With
                            
                            With .Fields(2)
                                .Value = "Pago de compra"                                               'tipo
                            End With
                            
                            If Val(Text2(1)) >= Val(Text2(0)) Then
                                With .Fields(3)
                                    .Value = Replace(Format(Val(Text2(0)) * -1, "0.00"), ",", ".")      'cantidad
                                End With
                            Else
                                With .Fields(3)
                                    .Value = Replace(Format(Val(Text2(1)) * -1, "0.00"), ",", ".")      'cantidad
                                End With
                            End If
                            
                            With .Fields(4)
                                .Value = Text1(0)                                                       'folio
                            End With
                            
                            With .Fields(5)
                                .Value = "No"                                                           'cancelado
                            End With
                            
                            With .Fields(6)
                                .Value = frmMenuInicial.Combo1.Text                                     'caja
                            End With
                            
                            With .Fields("created_by")
                                .Value = StUsuario                                                      'usuario
                            End With
                                                    
                            With .Fields("creation_date")
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                            End With
                                                    
                            With .Fields("last_updated_by")
                                .Value = StUsuario                                                      'usuario
                            End With
                                                    
                            With .Fields("last_update_date")
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                            End With
                        .Update
                        .Requery
                    End With
                End If
            End With
            'asignar valores a campos cabecera
            With Text1(0)
                v3 = .Text              'folio
            End With
            v6 = Date                   'fecha
            v18 = "No"                  'cancelado"
            With Text1(3)
                v19 = .Text             'comentarios
            End With
            v20 = StTipoVentasCompras   'tipo
            v21 = "C" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdTransaccion 'lote
            With Text2(1)
                DineroRestante = Val(.Text)
            End With
            
            With List1
                For i = 0 To .ListCount - 1
                    .ListIndex = i
                    'asignar valores a campos lineas
                    v8 = Trim(Mid(.Text, 1, 10))                                                    'idarticulo
                    v7 = Get_ItemTipo(v8)                                                           'Tipoarticulo
                    v9 = Get_ItemCod(v8)                                                            'codigo articulo
                    v10 = Get_ItemDesc(v8)                                                          'descripcion articulo
                    v11 = Replace(Format(Val(Trim(Mid(.Text, 60, 15))), "0.00"), ",", ".")          'cantidad
                    v12 = Get_ItemUDM(v8)                                                           'UDM
                    v13 = Replace(Format(Val(Trim(Mid(.Text, 76, 15))), "0.00"), ",", ".")          'precio
                    v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")                    'subtotal
                    viva = Get_ItemIva(v8)
                    v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")        'iva
                    v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                    'total
                    ControlLote = Get_ItemLote(v8)
                    vCategoria = Get_ItemCategoria(v8)
                    If DineroRestante = 0 Then
                        v17 = "0"
                    Else
                        If DineroRestante >= Val(v16) Then
                            v17 = v16
                            DineroRestante = DineroRestante - Val(v16)
                        Else
                            v17 = Replace(Format(DineroRestante, "0.00"), ",", ".")
                            DineroRestante = 0
                        End If
                    End If                                                                          'totalpagado
                    
                    With Rs5
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        If v12 <> "Servicio" And vCategoria = "Inventario" Then
                            .AddNew
                                With .Fields(1)
                                    .Value = v8                                                             'id
                                End With
                                
                                With .Fields(2)
                                    .Value = v9                                                             'codigo
                                End With
                                
                                With .Fields(3)
                                    .Value = v10                                                            'descripcion
                                End With
                                
                                With .Fields(4)
                                    .Value = Date                                                           'fecha
                                End With
                                
                                With .Fields(5)
                                    .Value = "Recepción de compra"                                          'tipo de treansaccion
                                End With
                                
                                With .Fields(6)
                                    .Value = Replace(Format(Val(v11), "0.00"), ",", ".")                    'cantidad
                                End With
                                
                                With .Fields(7)
                                    .Value = v12                                                            'udm
                                End With
                                
                                With .Fields(8)
                                    .Value = v3                                                             'folio
                                End With
                                
                                With .Fields(9)
                                    .Value = v18                                                            'cancelado
                                End With
                                
                                If ControlLote = True Then
                                    With .Fields(10)
                                        .Value = v21                                                        'lote
                                    End With
                                End If
                            
                                With .Fields("created_by")
                                    .Value = StUsuario                                                      'usuario
                                End With
                                                        
                                With .Fields("creation_date")
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                End With
                                                        
                                With .Fields("last_updated_by")
                                    .Value = StUsuario                                                      'usuario
                                End With
                                                        
                                With .Fields("last_update_date")
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                End With
                            .Update
                            .Requery
                        End If
                        .Close
                    End With
                    'guardar compra o venta
                    With Rs3
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                            With .Fields(1)
                                .Value = v1                                                             'id clienteproveedor
                            End With
                            
                            With .Fields(2)
                                .Value = v2                                                             'nombre cliente proveedor
                            End With
                            
                            With .Fields(3)
                                .Value = v3                                                             'folio
                            End With
                            
                            With .Fields(5)
                                .Value = v6                                                             'fecha
                            End With
                            
                            With .Fields(6)
                                .Value = v7                                                             'Tipoarticulo
                            End With
                            
                            With .Fields(7)
                                .Value = v8                                                             'idarticulo
                            End With
                            
                            With .Fields(8)
                                .Value = v9                                                             'codigo articulo
                            End With
                            
                            With .Fields(9)
                                .Value = v10                                                            'descripcion articulo
                            End With
                            
                            With .Fields(10)
                                .Value = Replace(Format(Val(v11), "0.00"), ",", ".")                    'cantidad
                            End With
                            
                            With .Fields(11)
                                .Value = v12                                                            'UDM
                            End With
                            
                            With .Fields(12)
                                .Value = v13                                                            'precio
                            End With
                            
                            With .Fields(13)
                                .Value = v14                                                            'subtotal
                            End With
                            
                            With .Fields(14)
                                .Value = v15                                                            'iva
                            End With
                            
                            With .Fields(15)
                                .Value = v16                                                            'total
                            End With
                            
                            With .Fields(16)
                                .Value = v17                                                            'totalpagado
                            End With
                            
                            With .Fields(17)
                                .Value = v18                                                            'cancelado
                            End With
                            
                            With .Fields(18)
                                .Value = v19                                                            'comentarios
                            End With
                            
                            With .Fields(19)
                                .Value = v20                                                            'tipo
                            End With
                            
                            With .Fields(20)
                                .Value = Replace(Replace(v3, "V-", ""), "C-", "")                       'NUM_FOLIO
                            End With
                            
                            If ControlLote = True Then
                                With .Fields(21)
                                    .Value = v21                                                        'lote
                                End With
                            End If
                            
                            With .Fields("created_by")
                                .Value = StUsuario                                                      'usuario
                            End With
                                                    
                            With .Fields("creation_date")
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                            End With
                                                    
                            With .Fields("last_updated_by")
                                .Value = StUsuario                                                      'usuario
                            End With
                                                    
                            With .Fields("last_update_date")
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                            End With
                        .Update
                        .Requery
                        .Close
                    End With
                    'guardar lote
                    InLoteExiste = Get_LoteExiste(v21, v8)
                    If ControlLote = True And InLoteExiste = 0 Then
                        With Rs7
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                With .Fields(1)
                                    .Value = v8                                                             'idarticulo
                                End With
                                
                                With .Fields(2)
                                    .Value = v21                                                            'lote
                                End With
                                
                                With .Fields(3)
                                    .Value = "Compras"                                                      'tipo
                                End With
                            
                                With .Fields("created_by")
                                    .Value = StUsuario                                                      'usuario
                                End With
                                                        
                                With .Fields("creation_date")
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                End With
                                                        
                                With .Fields("last_updated_by")
                                    .Value = StUsuario                                                      'usuario
                                End With
                                                        
                                With .Fields("last_update_date")
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                End With
                            .Update
                            .Requery
                            .Close
                        End With
                    End If
                Next i
            End With
            'imprimir ticket
            With Rs6
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from PO_TRANSACTION_TICKET where folio = '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                If .RecordCount <> 0 Then
                    Unload dsrComprasVentas
                    With dsrComprasVentas
                        With Rs6
                            vTicketSubtotal = Get_SumSubtotal(.Fields(6))
                            vTicketIva = Get_SumIva(.Fields(6))
                            vTicketTotal = Get_SumTotal(.Fields(6))
                        End With
                        vTicketSubtotal = Replace(Format(Val(vTicketSubtotal), "0.00"), ",", ".")
                        vTicketIva = Replace(Format(Val(vTicketIva), "0.00"), ",", ".")
                        vTicketTotal = Replace(Format(Val(vTicketTotal), "0.00"), ",", ".")
                        Set .DataSource = Rs6
                        
                        With .Sections("Sección4")
                            With .Controls("Etiqueta2")
                                .Caption = "TICKET DE COMPRA"
                            End With
                            
                            With .Controls("Etiqueta30")
                                .Caption = "Usuario: " & StUsuario
                            End With
                            
                            With .Controls("Etiqueta3")
                                .Caption = PcNombreEmpresa
                            End With
                            
                            With .Controls("Etiqueta4")
                                .Caption = PcRFC
                            End With
                            
                            With .Controls("Etiqueta5")
                                .Caption = PcDireccion
                            End With
                            
                            With .Controls("Etiqueta6")
                                .Caption = PcTelefono
                            End With
                            
                            With .Controls("Etiqueta11")
                                .Caption = Rs6.Fields(2) 'cliente
                            End With
                            
                            With .Controls("Etiqueta12")
                                .Caption = Rs6.Fields(3) 'calle
                            End With
                            
                            With .Controls("Etiqueta13")
                                .Caption = Rs6.Fields(4) 'colonia
                            End With
                            
                            With .Controls("Etiqueta14")
                                .Caption = Rs6.Fields(5) 'telefono
                            End With
                            
                            With .Controls("Etiqueta17")
                                .Caption = Rs6.Fields(7) 'fecha
                            End With
                            
                            With .Controls("Etiqueta18")
                                .Caption = Rs6.Fields(6) 'folio
                            End With
                        End With
                        
                        With .Sections("Sección1")
                            With .Controls("Texto1")
                                .DataField = "cantidad"
                            End With
                            
                            With .Controls("Texto2")
                                .DataField = "articulo"
                            End With
                            
                            With .Controls("Texto3")
                                .DataField = "subtotal"
                            End With
                        End With
                        
                        With .Sections("Sección5")
                            With .Controls("Etiqueta23")
                                .Caption = "$ " & vTicketSubtotal    'subtotal
                            End With
                            
                            With .Controls("Etiqueta26")
                                .Caption = "$ " & vTicketIva         'iva
                            End With
                            
                            With .Controls("Etiqueta27")
                                .Caption = "$ " & vTicketTotal       'total
                            End With
                            
                            With .Controls("Label1")
                                .Visible = False
                            End With
                            
                            With .Controls("Label2")
                                .Visible = False
                            End With
                            
                            With .Controls("Label3")
                                .Visible = False
                            End With
                            
                            With .Controls("Label4")
                                .Visible = False
                            End With
                            
                            With .Controls("Label5")
                                .Visible = False
                            End With
                            
                            With .Controls("Label6")
                                .Visible = False
                            End With
                            
                            With .Controls("Etiqueta25")
                                .Visible = False
                            End With
                            
                            With .Controls("Etiqueta28")
                                .Visible = False
                            End With
                        End With
                        .Show 1
                    End With
                End If
                .Close
            End With
            Unload frmCompras
            Set frmCompras = Nothing
            
            With frmCompras
                .Show
            End With
        Else
            Exit Sub
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        Unload dsrComprasVentas
        With Rs
            If .State = 1 Then .Close
        End With
        
        With RS1
            If .State = 1 Then .Close
        End With
        
        With Rs2
            If .State = 1 Then .Close
        End With
        
        With Rs3
            If .State = 1 Then .Close
        End With
        
        With Rs4
            If .State = 1 Then .Close
        End With
        
        With Rs5
            If .State = 1 Then .Close
        End With
        
        With Rs6
            If .State = 1 Then .Close
        End With
        
        With Rs7
            If .State = 1 Then .Close
        End With
        
        With Rs9
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set RS1 = Nothing
        Set Rs2 = Nothing
        Set Rs3 = Nothing
        Set Rs4 = Nothing
        Set Rs5 = Nothing
        Set Rs6 = Nothing
        Set Rs7 = Nothing
        Set Rs9 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCompras:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
