VERSION 5.00
Begin VB.Form frmPedidos 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Nuevo Pedido"
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
      Height          =   11535
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17295
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8895
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   17175
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3960
            Width           =   1455
         End
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
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   14
            Top             =   8160
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   12
            Top             =   7200
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   13
            Top             =   7680
            Width           =   6495
         End
         Begin VB.TextBox Text1 
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
            Left            =   2400
            TabIndex        =   7
            Top             =   2640
            Width           =   14415
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
            Left            =   240
            TabIndex        =   11
            Top             =   5040
            Width           =   16575
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            MaxLength       =   7
            TabIndex        =   6
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            MaxLength       =   7
            TabIndex        =   5
            Top             =   1680
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
            Left            =   2400
            TabIndex        =   4
            Top             =   1140
            Width           =   14415
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
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
            Index           =   0
            Left            =   2400
            TabIndex        =   2
            Top             =   600
            Width           =   12615
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "BUSCAR"
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
            Index           =   2
            Left            =   15360
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Width           =   1455
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
            Index           =   2
            Left            =   2400
            TabIndex        =   8
            Top             =   3120
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmPedidos.frx":0000
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
            Left            =   360
            TabIndex        =   26
            Top             =   4680
            Width           =   12615
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
            Index           =   8
            Left            =   8040
            TabIndex        =   25
            Top             =   7200
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
            Index           =   12
            Left            =   8040
            TabIndex        =   24
            Top             =   7680
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
            Index           =   13
            Left            =   120
            TabIndex        =   23
            Top             =   2640
            Width           =   2055
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
            Index           =   15
            Left            =   8040
            TabIndex        =   22
            Top             =   8160
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO DE VENTA"
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
            Left            =   120
            TabIndex        =   21
            Top             =   3120
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
            Index           =   17
            Left            =   -480
            TabIndex        =   20
            Top             =   2160
            Width           =   2535
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
            Index           =   18
            Left            =   120
            TabIndex        =   19
            Top             =   1680
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
            Index           =   19
            Left            =   120
            TabIndex        =   18
            Top             =   1140
            Width           =   2055
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
            Index           =   20
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   2055
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
            Index           =   21
            Left            =   1200
            TabIndex        =   16
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Anadir 
         Caption         =   "Añadir Cliente"
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
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmPedidos
'Proposito:     Registra pedidos de venta
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        21/05/2021     Alfredo Hernandez    Se agrego usuario, fecha de creacion
'                                               y modificacion a todos los insert
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset    'folio
Dim RS1 As New adodb.Recordset    'clientesproveedores
Dim Rs2 As New adodb.Recordset    'items
Dim Rs3 As New adodb.Recordset    'ventascompras
Dim Rs6 As New adodb.Recordset    'ticket
Dim Rs13 As New adodb.Recordset    'cabecera
'//OTROS
Dim TipoErr As Long
Dim i As Long
Dim X As Long
Dim intX As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim c4 As Long
Dim Prt As Printer
'//VALORES PARA INSERTAR
Dim v1 As Long    'idclienteproveedor
Dim v2 As String    'nombre cliente proveedor
Dim v3 As String    'folio
Dim v4 As String    'LugarVenta
Dim v5 As String    'mesa
Dim v6 As Date    'fecha
Dim v7 As String    'Tipoarticulo
Dim v8 As Long    'idarticulo
Dim v9 As String    'codigo articulo
Dim v10 As String    'descripcion articulo
Dim v11 As String    'cantidad
Dim v12 As String    'UDM
Dim v13 As String    'precio
Dim v14 As String    'subtotal
Dim v15 As String    'iva
Dim v16 As String    'total
Dim v17 As String    'totalpagado
Dim v18 As String    'cancelado
Dim v19 As String    'comentarios
Dim v20 As String    'tipo
Dim IdTransaccion As Long           'folio
'//ARTICULOS
Dim InItemId As Long
Dim vCategoria As String
'//CLIENTES
Dim Credito As String   'credito del cliente
Dim CreditoUsado As String        'Credito usado del cliente
Dim DiasCredito As Long         'Dias de credito del cliente
Dim DiasCreditoUsado As Long              'Dias de credito usados por el cliente
Dim ClienteMayorista As String            '¿Es cliente mayorista?
Dim ListaPrecios As Long          'lista de precios cliente
'//VENTAS
Dim listSubtotal As String
Dim listIva As String
Dim listTotal As String
Dim vLstCantidad As String
Dim vLstPrecio As String
Dim viva As String
Dim viid As String
Dim videscripcion As String
Dim vicantidad As String
Dim viprecio As String
Dim Array_Comentarios() As String

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
        .Open "SELECT 1+ Max(num_folio) AS IdVenta From PO_LINES_ALL WHERE tipo = 'Pedidos'", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        .MoveFirst
        If IsNull(Rs!IdVenta) = False Then
            IdTransaccion = Rs!IdVenta
        Else
            IdTransaccion = 1
        End If

        With Text1(0)
            .Text = "P-" & IdTransaccion
        End With
    End With

    For i = 1 To 3
        With Text1(i)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    For i = 0 To 2
        With Combo1(i)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    With Label1(6)
        .Visible = True
    End With

    With Combo1(2)
        .Visible = True
    End With

    With RS1
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from HZ_PARTY where cliente = 'Si' order by 2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
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

    With Combo1(2)
        .AddItem "Local"
        .AddItem "Domicilio"
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
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
            Exit Sub
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
                    ListaPrecios = 0
                End With
            Else
                .BackColor = COLOR_NORMAL
                With RS1
                    .Filter = ""
                    .Requery
                    .Filter = "nombre like '" & Combo1(0) & "'"
                    .Requery
                    If .RecordCount <> 0 Then v1 = .Fields(0).Value

                    If .RecordCount <> 0 Then v2 = .Fields(1).Value

                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    Case 1
        If ListaPrecios = 0 Then
            MsgBox "Seleccionar Cliente primero", vbOKOnly, "Información"
        Else
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = COLOR_NO_ENCONTRADO
                    With Rs2
                        .Filter = ""
                        .Requery
                    End With
                    Text1(2).Text = ""
                Else
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(7).Value, ",", ".")
                        End If
                    End With
                End If
            End With
        End If
    Case 2
        With Combo1(2)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
                With Combo1(3)
                    .Enabled = False
                End With
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
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
                    ListaPrecios = 0
                End With
            Else
                ' Backcolor normal cuando hay coincidencia
                .BackColor = COLOR_NORMAL
                With RS1
                    .Filter = ""
                    .Filter = "nombre like '" & Combo1(0) & "'"
                    .Requery
                    If .RecordCount <> 0 Then v1 = .Fields(0)

                    If .RecordCount <> 0 Then v2 = .Fields(1)

                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    Case 1
        If ListaPrecios = 0 Then
            MsgBox "Seleccionar Cliente primero", vbOKOnly, "Información"
        Else
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
                    Text1(2).Text = ""
                Else
                    ' Backcolor normal cuando hay coincidencia
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(7).Value, ",", ".")
                        End If
                    End With
                End If
            End With
        End If
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_Change(Index As Integer)
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
                    ListaPrecios = 0
                End With
            Else
                .BackColor = COLOR_NORMAL
                With RS1
                    .Filter = ""
                    .Filter = "nombre like '" & Combo1(0) & "'"
                    .Requery
                    If .RecordCount <> 0 Then v1 = .Fields(0)

                    If .RecordCount <> 0 Then v2 = .Fields(1)

                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    Case 1
        If ListaPrecios = 0 Then
            MsgBox "Seleccionar Cliente primero", vbOKOnly, "Información"
        Else
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = COLOR_NO_ENCONTRADO
                    With Rs2
                        .Filter = ""
                        .Requery
                    End With
                    Text1(2).Text = ""
                Else
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(7).Value, ",", ".")
                        End If
                    End With
                End If
            End With
        End If
    Case 2
        With Combo1(2)
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Combo1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Text1(1)
            If .Text = "" Or Val(.Text) <= 0 Then
                MsgBox "Cantidad no válida", vbCritical, "Error"
                Exit Sub
            End If

            If Combo1(1) <> "" And .Text <> "" And Text1(2) <> "" Then
                With Rs2
                    viid = .Fields(0).Value
                End With

                With Combo1(1)
                    videscripcion = Mid(.Text, 1, 47)
                End With
                vicantidad = Replace(Format(Val(.Text), "0.00"), ",", ".")
                If ListaPrecios = 1 And Val(.Text) >= 5 And Rs2.RecordCount <> 0 Then
                    With Rs2
                        viprecio = Replace(Format(.Fields(4).Value, "0.00"), ",", ".")
                    End With
                Else
                    With Text1(2)
                        viprecio = Replace(Format(Val(.Text), "0.00"), ",", ".")
                    End With
                End If
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
                .Text = ""
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
        End With
    Case 1
        With List1
            intX = .ListIndex
            .RemoveItem intX
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
    Case 2
        TipoBusquedaCliente = "Pedido"
        With frmBuscadorClientes
            .Show 1
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Command1_Click" & vbTab & err.Number & vbTab & err.Description
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Anadir_Click()
    On Error GoTo errHandler
    InTipoAltaClienteProveedor = 1
    StTipoClienteProveedor = "Cliente"
    With frmClientesNuevo
        .Caption = "Añadir nuevo Cliente"
        .Show 1
    End With
    Unload frmPedidos
    Set frmPedidos = Nothing

    With frmPedidos
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Anadir_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    With List1
        If .ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
            With Combo1(2)
                If .Text = "" Then
                    MsgBox "Llenar el lugar de la venta", vbOKOnly, "Advertencia"
                    Exit Sub
                End If

                'confirmacion datos cliente
                If .Text = "Domicilio" Then
                    IdCliente = v1
                    With frmClientesConfirmacionDatos
                        .Show 1
                    End With
                End If
            End With

            '=======================================================================================
            'I  N   I   C   I   O
            '=======================================================================================
            'Actualizar Folio
            With Rs
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "SELECT 1+ Max(num_folio) AS IdVenta From PO_LINES_ALL WHERE tipo = 'Pedidos'", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                .MoveFirst
                If IsNull(Rs!IdVenta) = False Then
                    IdTransaccion = Rs!IdVenta
                Else
                    IdTransaccion = 1
                End If

                With Text1(0)
                    .Text = "P-" & IdTransaccion
                End With
            End With

            With Text1(0)
                v3 = .Text              'folio
            End With
            v6 = Date                   'fecha
            v18 = "No"                  'cancelado"
            With Text1(3)
                v19 = .Text             'comentarios
            End With
            v20 = StTipoVentasCompras   'tipo
            With Combo1(2)
                v4 = .Text              'LugarVenta
            End With

            For i = 0 To .ListCount - 1
                .ListIndex = i
                'asignar valores a campos lineas
                v8 = Trim(Mid(.Text, 1, 10))                                                'idarticulo
                v7 = Get_ItemTipo(v8)                                                       'Tipoarticulo
                v9 = Get_ItemCod(v8)                                                        'codigo articulo
                v10 = Get_ItemDesc(v8)                                                      'descripcion articulo
                v11 = Replace(Format(Val(Trim(Mid(.Text, 60, 15))), "0.00"), ",", ".")      'cantidad
                v12 = Get_ItemUDM(v8)                                                       'UDM
                v13 = Replace(Format(Val(Trim(Mid(.Text, 76, 15))), "0.00"), ",", ".")      'precio
                viva = Get_ItemIva(v8)
                vCategoria = Get_ItemCategoria(v8)
                v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")                'subtotal
                v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")    'iva
                v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                'total
                'guardar compra o venta
                With Rs3
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                    With .Fields(1)
                        .Value = v1                                                             'idclienteproveedor
                    End With

                    With .Fields(2)
                        .Value = v2                                                             'nombre cliente proveedor
                    End With

                    With .Fields(3)
                        .Value = v3                                                             'folio
                    End With

                    With .Fields(4)
                        .Value = v4                                                             'LugarVenta
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
                        .Value = v11                                                            'cantidad
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
                        .Value = 0                                                              'totalpagado
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
                        .Value = Replace(v3, "P-", "")                                          'NUM_FOLIO
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
            Next i
            'imprimir ticket
            With Rs6
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from PO_TRANSACTION_TICKET where folio = '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                If .RecordCount <> 0 Then
                    Unload dsrPedidoVenta
                    With dsrPedidoVenta
                        Set .DataSource = Rs6
                        With .Sections("Sección2")
                            With .Controls("Label1")
                                .Caption = "Folio - " & v3
                            End With

                            With .Controls("Label2")
                                .Caption = "Comentarios: " & v19
                            End With
                        End With
                        With .Sections("Sección1")
                            With .Controls("Texto1")
                                .DataField = "cantidad"
                            End With

                            With .Controls("Texto2")
                                .DataField = "articulo"
                            End With
                        End With
                        .Show 1
                    End With
                    Unload dsrComprasVentas
                End If
                .Close
            End With
            Unload frmPedidos
            Set frmPedidos = Nothing

            With frmPedidos
                .Show 1
            End With

            Exit Sub
            '======================================================================================
            'F   I   N
            '======================================================================================
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Advertencia"
            Exit Sub
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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

    With Rs6
        If .State = 1 Then .Close
    End With

    With Rs13
        If .State = 1 Then .Close
    End With

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set RS1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Rs6 = Nothing
    Set Rs13 = Nothing
    Set Cn = Nothing
    Set frmPedidos = Nothing
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidos:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
