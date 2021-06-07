VERSION 5.00
Begin VB.Form frmAgregarArticuloVentas 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Articulo a Ventas"
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
      Height          =   9015
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   17295
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8895
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   0
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
            TabIndex        =   7
            Top             =   3720
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
            TabIndex        =   8
            Top             =   3720
            Width           =   1455
         End
         Begin VB.TextBox Text1 
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
            Index           =   8
            Left            =   2400
            TabIndex        =   6
            Top             =   3040
            Width           =   3015
         End
         Begin VB.TextBox Text1 
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
            Index           =   7
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   600
            Width           =   14415
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
            TabIndex        =   12
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   5
            Top             =   2560
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
            Height          =   2340
            Left            =   240
            TabIndex        =   9
            Top             =   4680
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
            TabIndex        =   4
            Top             =   2080
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
            TabIndex        =   3
            Top             =   1600
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
            TabIndex        =   2
            Top             =   1080
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
            TabIndex        =   0
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAgregarArticuloVentas.frx":0000
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
            TabIndex        =   25
            Top             =   4320
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
            Index           =   10
            Left            =   8040
            TabIndex        =   24
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
            Index           =   7
            Left            =   8040
            TabIndex        =   23
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
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   2560
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
            Index           =   9
            Left            =   8040
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   3040
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
            TabIndex        =   19
            Top             =   2080
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
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   1600
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
            Left            =   120
            TabIndex        =   17
            Top             =   1080
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
            Index           =   1
            Left            =   120
            TabIndex        =   16
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
            Index           =   0
            Left            =   1200
            TabIndex        =   15
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
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
Attribute VB_Name = "frmAgregarArticuloVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmAgregarArticuloVentas
'Proposito:     Agregar un articulo a una venta registrada previamente
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        19/05/2021     Alfredo Hernandez    Se agrego usuario, fecha de creacion
'                                               y modificacion a todos los insert
'
'1.2        19/05/2021     Alfredo Hernandez    Se agrego validacion para inv.
'                                               negativos
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim RS1 As New adodb.Recordset    'clientesproveedores
Dim Rs2 As New adodb.Recordset    'items
Dim Rs3 As New adodb.Recordset    'ventascompras
Dim Rs4 As New adodb.Recordset    'lista de ingredientes
Dim Rs5 As New adodb.Recordset    'movimientos de inventarios
Dim Rs6 As New adodb.Recordset    'ticket
Dim Rs9 As New adodb.Recordset    'entrada de dinero
Dim Rs12 As New adodb.Recordset    'lotes
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
'//VENTAS
Dim InItemId As Long
Dim ListaPrecios As Long          'lista de precios cliente
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
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim c4 As Long
Dim vCategoria As String
Dim vDevExiste As Long
Dim vCantidadDev As String
'//LOTE
Dim ControlLote As Boolean
Dim CantidadRestante As String
Dim vLote As String
Dim vCantidadLote As String
Dim vCurrentLote As String
Dim InLoteExiste As Long
'//OTROS
Dim i As Long
Dim X As Long
Dim intX As Long
Dim v As Long

Private Sub Form_Load()
    On Error GoTo errHandler
    For i = 1 To (3)
        With Text1(i)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    With Combo1(1)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Label1(6)
        .Visible = True
    End With

    With Text1(8)
        .Visible = True
    End With

    With RS1
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from HZ_PARTY where cliente = 'Si' order by 2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Filter = ""
        .Requery
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
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Form_Load" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
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

                    With Text1(2)
                        .Text = ""
                    End With
                Else
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            .Text = Replace(Rs2.Fields(7).Value, ",", ".")
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    Static cadena As String
    Select Case Index
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

                    With Text1(2)
                        .Text = ""
                    End With
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
                            .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            .Text = Replace(Rs2.Fields(7).Value, ",", ".")
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
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

            If ListaPrecios = 1 And Val(Text1(1).Text) >= 5 And Rs2.RecordCount <> 0 Then
                With Rs2
                    viprecio = Replace(Format(.Fields(4).Value, "0.00"), ",", ".")
                End With
            Else
                With Text1(2)
                    viprecio = Replace(Format(Val(.Text), "0.00"), ",", ".")
                End With
            End If
            With Rs2
                If Get_ItemUDM(.Fields(0).Value) <> "Servicio" And Get_ItemCategoria(.Fields(0).Value) = "Inventario" Then
                    If PcInventarios = False Then
                        If Val(Get_CantidadItem(.Fields(0).Value)) < Val(vicantidad) Then
                            MsgBox "Existencia insuficiente, no se puede agregar a la venta", vbCritical, "Advertencia"
                            Exit Sub
                        End If
                    End If
                End If
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
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Command1_Click" & vbTab & err.Number & vbTab & err.Description
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
    Case 7
        With Text1(7)
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
                    .Filter = "nombre like '" & Text1(7) & "'"
                    .Requery
                    v1 = .Fields(0)
                    v2 = .Fields(1)
                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With List1
            If .ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
                'asignar valores a campos cabecera
                v3 = Text1(0)    'folio
                v6 = Date    'fecha
                v18 = "No"    'cancelado"
                v19 = Text1(3)    'comentarios
                v20 = StTipoVentasCompras    'tipo
                v4 = Text1(8)    'LugarVenta
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
                    viva = Get_ItemIva(v8)
                    v17 = 0                                                                         'totalpagado
                    'Inventario o Gasto
                    vCategoria = Get_ItemCategoria(v8)
                    'lote
                    ControlLote = Get_ItemLote(v8)
                    'Devoluciones
                    If Val(v11) < 0 Then
                        vDevExiste = Get_DevItemExiste(v8, v3)
                        If vDevExiste = 0 Then
                            MsgBox "No se puede devolver el artículo porque no existe en la venta", vbCritical, "Error"
                        Else
                            'si tiene control de lote
                            If ControlLote = True Then
                                CantidadRestante = v11
                                'mientras no se complete la cantidad necesaria
                                While Val(CantidadRestante) < 0
                                    'obtenemos lote mas antiguo y existencia de ese lote
                                    vLote = ""
                                    vLote = Get_LoteConsumoDev(v8, v3)
                                    vCantidadLote = Get_LoteConsumoCantidadDev(v8, v3)
                                    'si existe algun lote
                                    If vLote <> "" Then
                                        If Val(vCantidadLote) * -1 < Val(CantidadRestante) Then
                                            If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                                With Rs5
                                                    If .State = 1 Then .Close
                                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                                    .Requery
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
                                                        .Value = "Devolución de venta"                                          'tipo de treansaccion
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

                                                    With .Fields(10)
                                                        .Value = vLote                                                          'lote
                                                    End With

                                                    With .Fields(6)
                                                        .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
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
                                            v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")               'subtotal
                                            v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                                            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
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
                                                    .Value = CantidadRestante                                               'cantidad
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

                                                With .Fields(21)
                                                    .Value = vLote                                                          'lote
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
                                            CantidadRestante = "0"
                                        Else
                                            If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                                With Rs5
                                                    If .State = 1 Then .Close
                                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                                    .Requery
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
                                                        .Value = "Devolución de venta"                                          'tipo de treansaccion
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

                                                    With .Fields(10)
                                                        .Value = vLote                                                          'lote
                                                    End With

                                                    With .Fields(6)
                                                        .Value = Replace(Format(Val(vCantidadLote), "0.00"), ",", ".")          'cantidad
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
                                            v14 = Replace(Format(Val(vCantidadLote) * Val(v13) * -1, "0.00"), ",", ".")             'subtotal
                                            v15 = Replace(Format(Val(vCantidadLote) * Val(v13) * Val(viva) * -1, "0.00"), ",", ".")    'iva
                                            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
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
                                                    .Value = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".")    'cantidad
                                                End With

                                                With .Fields(11)
                                                    .Value = v12                                                           'UDM
                                                End With

                                                With .Fields(12)
                                                    .Value = v13                                                           'precio
                                                End With

                                                With .Fields(13)
                                                    .Value = v14                                                           'subtotal
                                                End With

                                                With .Fields(14)
                                                    .Value = v15                                                           'iva
                                                End With

                                                With .Fields(15)
                                                    .Value = v16                                                           'total
                                                End With

                                                With .Fields(16)
                                                    .Value = v17                                                           'totalpagado
                                                End With

                                                With .Fields(17)
                                                    .Value = v18                                                           'cancelado
                                                End With

                                                With .Fields(18)
                                                    .Value = v19                                                           'comentarios
                                                End With

                                                With .Fields(19)
                                                    .Value = v20                                                           'tipo
                                                End With

                                                With .Fields(20)
                                                    .Value = Replace(Replace(v3, "V-", ""), "C-", "")                      'NUM_FOLIO
                                                End With

                                                With .Fields(21)
                                                    .Value = vLote                                                         'lote
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
                                            CantidadRestante = Replace(Format(Val(CantidadRestante) + Val(vCantidadLote), "0.00"), ",", ".")
                                        End If
                                        'si no existen lotes no se devuelve
                                    Else
                                        MsgBox "No se puede devolver " & CantidadRestante & " porque no existen lotes para la devolución", vbCritical, "Advertencia"
                                        CantidadRestante = "0"
                                    End If
                                Wend
                            Else
                                v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")                'subtotal
                                v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")    'iva
                                v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                'total
                                vCantidadDev = Get_CantidadDev(v8, v3)
                                If Val(vCantidadDev) > (Val(v11) * -1) Then
                                    If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                        'salida por venta
                                        With Rs5
                                            If .State = 1 Then .Close
                                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                            .Requery
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
                                                .Value = "Devolución de venta"                                          'tipo de treansaccion
                                            End With

                                            With .Fields(6)
                                                .Value = Replace(Format(Val(v11) * -1, "0.00"), ",", ".")               'cantidad
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
                                    Rs5.Close
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
                                Else
                                    MsgBox "No se puede devolver " & Val(v11) * -1 & " porque excede la cantidad vendida que es de " & Val(vCantidadDev), vbCritical, "Advertencia"
                                End If
                            End If
                        End If
                    End If
                    'venta normal
                    v = 0    ' Ticket de pedido
                    If Val(v11) > 0 Then
                        'si tiene control de lote
                        If ControlLote = True Then
                            CantidadRestante = v11
                            vCurrentLote = "V" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdTransaccion
                            'mientras no se complete la cantidad necesaria
                            While Val(CantidadRestante) > 0
                                'obtenemos lote mas antiguo y existencia de ese lote
                                vLote = ""
                                vLote = Get_LoteConsumo(v8)
                                vCantidadLote = Get_LoteConsumoCantidad(v8)
                                'si existe algun lote
                                If vLote <> "" Then
                                    If Val(vCantidadLote) > Val(CantidadRestante) Then
                                        If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                            With Rs5
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
                                                    .Value = "Salida por venta"                                             'tipo de treansaccion
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

                                                With .Fields(10)
                                                    .Value = vLote                                                          'lote
                                                End With

                                                With .Fields(6)
                                                    .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
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
                                        v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")               'subtotal
                                        v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                                        v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
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
                                                .Value = CantidadRestante                                               'cantidad
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

                                            With .Fields(21)
                                                .Value = vLote                                                          'lote
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
                                        CantidadRestante = "0"
                                        v = v + 1
                                    Else
                                        If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                            With Rs5
                                                If .State = 1 Then .Close
                                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                                .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                                .Requery
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
                                                    .Value = "Salida por venta"                                             'tipo de treansaccion
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

                                                With .Fields(10)
                                                    .Value = vLote                                                          'lote
                                                End With

                                                With .Fields(6)
                                                    .Value = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".")     'cantidad
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
                                        v14 = Replace(Format(Val(vCantidadLote) * Val(v13), "0.00"), ",", ".")              'subtotal
                                        v15 = Replace(Format(Val(vCantidadLote) * Val(v13) * Val(viva), "0.00"), ",", ".")  'iva
                                        v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                        'total
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
                                                .Value = vCantidadLote                                                  'cantidad
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

                                            With .Fields(21)
                                                .Value = vLote                                                          'lote
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
                                        CantidadRestante = Val(CantidadRestante) - Val(vCantidadLote)
                                        v = v + 1
                                    End If
                                    'si no existen lotes creamos uno
                                Else
                                    InLoteExiste = Get_LoteExiste(vCurrentLote, v8)
                                    If InLoteExiste = 0 Then
                                        With Rs12
                                            If .State = 1 Then .Close
                                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                            .Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                            .Requery
                                            .AddNew
                                            With .Fields(1)
                                                .Value = v8                                                             'idarticulo
                                            End With

                                            With .Fields(2)
                                                .Value = vCurrentLote                                                   'lote
                                            End With

                                            With .Fields(3)
                                                .Value = "Venta"                                                        'tipo"
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

                                    If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                        With Rs5
                                            If .State = 1 Then .Close
                                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                            .Requery
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
                                                .Value = "Salida por venta"                                             'tipo de treansaccion
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

                                            With .Fields(10)
                                                .Value = vCurrentLote                                                   'lote
                                            End With

                                            With .Fields(6)
                                                .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
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
                                    v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")               'subtotal
                                    v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                                    v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
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
                                            .Value = CantidadRestante                                               'cantidad
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

                                        With .Fields(21)
                                            .Value = vCurrentLote                                                   'lote
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
                                    CantidadRestante = "0"
                                    v = v + 1
                                End If
                            Wend
                        Else
                            v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")                'subtotal
                            v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")    'iva
                            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                'total
                            With Rs5
                                If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                    '   salida por venta
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
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
                                        .Value = "Salida por venta"                                             'tipo de treansaccion
                                    End With

                                    With .Fields(6)
                                        .Value = Replace(Format(Val(v11) * -1, "0.00"), ",", ".")               'cantidad
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
                            v = v + 1
                        End If
                    End If
                Next i
                'imprimir ticket
                With Rs6
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select Top " & v & " * from PO_TRANSACTION_TICKET_P where folio = '" & Text1(0) & "' order by id desc;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount <> 0 Then
                        Unload dsrPedidoVenta
                        With dsrPedidoVenta
                            Set .DataSource = Rs6

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
                    End If
                End With
                With frmHistorialVentas
                    .Enabled = True
                End With
                Unload Me
            Else
                MsgBox "Llenar todos los campos", vbCritical, "Advertencia"
                Exit Sub
            End If
        End With
    Else
        Exit Sub
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Salir_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
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

    With Rs9
        If .State = 1 Then .Close
    End With

    With Rs12
        If .State = 1 Then .Close
    End With

    Set RS1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Rs4 = Nothing
    Set Rs5 = Nothing
    Set Rs6 = Nothing
    Set Rs9 = Nothing
    Set Rs12 = Nothing
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAgregarArticuloVentas:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
