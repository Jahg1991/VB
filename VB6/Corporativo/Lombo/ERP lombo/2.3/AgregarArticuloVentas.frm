VERSION 5.00
Begin VB.Form frmAgregarArticuloVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Articulo a movimientos existentes"
   ClientHeight    =   8310
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13815
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
   ScaleHeight     =   8310
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7815
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   13335
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   8
            Left            =   1800
            TabIndex        =   6
            Top             =   2520
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   7
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   600
            Width           =   11295
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   6600
            TabIndex        =   24
            Top             =   7080
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   420
            Index           =   6
            Left            =   6600
            TabIndex        =   21
            Top             =   6120
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   420
            Index           =   5
            Left            =   6600
            TabIndex        =   20
            Top             =   6600
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   3
            Left            =   1800
            TabIndex        =   5
            Top             =   2040
            Width           =   11295
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2025
            Left            =   240
            TabIndex        =   9
            Top             =   3960
            Width           =   12855
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   1800
            Picture         =   "AgregarArticuloVentas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "AgregarArticuloVentas.frx":08D1
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   6000
            MaxLength       =   7
            TabIndex        =   4
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   3
            Top             =   1560
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   1800
            TabIndex        =   2
            Top             =   1080
            Width           =   11295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   10080
            TabIndex        =   0
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"AgregarArticuloVentas.frx":1105
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   360
            TabIndex        =   25
            Top             =   3600
            Width           =   12615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   10
            Left            =   4440
            TabIndex        =   23
            Top             =   6120
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Iva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   4440
            TabIndex        =   22
            Top             =   6600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   -360
            TabIndex        =   19
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   4440
            TabIndex        =   18
            Top             =   7080
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   -360
            TabIndex        =   17
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   3360
            TabIndex        =   16
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   -360
            TabIndex        =   15
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Artículo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   -360
            TabIndex        =   14
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   -360
            TabIndex        =   13
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Folio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   9000
            TabIndex        =   12
            Top             =   240
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
    Option Explicit
    
    '//RECORDSET
    Dim Rs1                 As New adodb.Recordset  'clientesproveedores
    Dim Rs2                 As New adodb.Recordset  'items
    Dim Rs3                 As New adodb.Recordset  'ventascompras
    Dim Rs4                 As New adodb.Recordset  'lista de ingredientes
    Dim Rs5                 As New adodb.Recordset  'movimientos de inventarios
    Dim Rs6                 As New adodb.Recordset  'ticket
    Dim Rs9                 As New adodb.Recordset  'entrada de dinero
    Dim Rs12                As New adodb.Recordset  'lotes
    
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
    Dim IdTransaccion       As Long                 'folio
    
    '//VENTAS
    Dim InItemId            As Long
    Dim ListaPrecios        As Long                 'lista de precios cliente
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
    Dim c1                  As Long
    Dim c2                  As Long
    Dim c3                  As Long
    Dim c4                  As Long
    Dim vCategoria          As String
    Dim vDevExiste          As Long
    Dim vCantidadDev        As String
    
    '//LOTE
    Dim ControlLote         As Long
    Dim CantidadRestante    As String
    Dim vLote               As String
    Dim vCantidadLote       As String
    Dim vCurrentLote        As String
    Dim InLoteExiste        As Long
    
    '//OTROS
    Dim i                   As Long
    Dim X                   As Long
    Dim intX                As Long
    Dim v                   As Long
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        For i = 1 To (3)
            Text1(i).BackColor = COLOR_NO_ENCONTRADO
        Next i
        
        Combo1(1).BackColor = COLOR_NO_ENCONTRADO
        
        Label1(6).Visible = True
        Text1(8).Visible = True
        
        With Rs1
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
                Combo1(1).Clear
                
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
                    viid = Rs2.Fields(0).Value
                    videscripcion = Mid(Combo1(1), 1, 47)
                    vicantidad = Replace(Format(Val(Text1(1)), "0.00"), ",", ".")
                    
                    If ListaPrecios = 1 And Val(Text1(1)) >= 5 And Rs2.RecordCount <> 0 Then
                        viprecio = Replace(Format(Rs2.Fields(4).Value, "0.00"), ",", ".")
                    Else
                        viprecio = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
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
                    
                    List1.AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                    
                    Text1(1) = ""
                    Text1(2) = ""
                    
                    With Combo1(1)
                        .Text = ""
                        .BackColor = COLOR_NO_ENCONTRADO
                        .SetFocus
                    End With
                    
                    listSubtotal = 0
                    listIva = 0
                    
                    For i = 0 To List1.ListCount - 1
                        List1.ListIndex = i
                        List1.SetFocus
                        
                        vLstCantidad = Trim(Mid(List1.Text, 60, 15))
                        vLstPrecio = Trim(Mid(List1.Text, 76, 15))
                        viva = Get_ItemIva(Trim(Mid(List1.Text, 1, 10)))
                        listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                        listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                    Next i
                    
                    listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                    listIva = Replace(Format(listIva, "0.00"), ",", ".")
                    listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                    listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                    
                    Text1(6) = listSubtotal
                    Text1(5) = listIva
                    Text1(4) = listTotal
                Else
                    MsgBox "Llenar todos los campos", vbCritical, "Error"
                    Combo1(1).SetFocus
                End If
                        
            Case 1
                With List1
                    intX = .ListIndex
                    .RemoveItem intX
                End With
                
                listSubtotal = 0
                listIva = 0
                
                For i = 0 To List1.ListCount - 1
                    List1.ListIndex = i
                    List1.SetFocus
                    vLstCantidad = Trim(Mid(List1.Text, 60, 15))
                    vLstPrecio = Trim(Mid(List1.Text, 76, 15))
                    viva = Get_ItemIva(Trim(Mid(List1.Text, 1, 10)))
                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                Next i
                
                listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                listIva = Replace(Format(listIva, "0.00"), ",", ".")
                listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                
                Text1(6) = listSubtotal
                Text1(5) = listIva
                Text1(4) = listTotal
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
                        
                        With Rs1
                            .Filter = ""
                            .Requery
                            
                            v1 = 0
                            v2 = ""
                            ListaPrecios = 0
                        End With
                    Else
                        .BackColor = COLOR_NORMAL
                        
                        With Rs1
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
            If List1.ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
                'asignar valores a campos cabecera
                v3 = Text1(0) 'folio
                v6 = Date 'fecha
                v18 = "No" 'cancelado"
                v19 = Text1(3) 'comentarios
                v20 = StTipoVentasCompras 'tipo
                v4 = Text1(8) 'LugarVenta
                
                For i = 0 To List1.ListCount - 1
                    List1.ListIndex = i
                    
                    'asignar valores a campos lineas
                    v8 = Trim(Mid(List1.Text, 1, 10))                                               'idarticulo
                    v7 = Get_ItemTipo(v8)                                                           'Tipoarticulo
                    v9 = Get_ItemCod(v8)                                                            'codigo articulo
                    v10 = Get_ItemDesc(v8)                                                          'descripcion articulo
                    v11 = Replace(Format(Val(Trim(Mid(List1.Text, 60, 15))), "0.00"), ",", ".")     'cantidad
                    v12 = Get_ItemUDM(v8)                                                           'UDM
                    v13 = Replace(Format(Val(Trim(Mid(List1.Text, 76, 15))), "0.00"), ",", ".")     'precio
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
                            If ControlLote = 1 Then
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
                                                        .Fields(1) = v8                                                             'id
                                                        .Fields(2) = v9                                                             'codigo
                                                        .Fields(3) = v10                                                            'descripcion
                                                        .Fields(4) = Date                                                           'fecha
                                                        .Fields(5) = "Devolución de venta"                                          'tipo de treansaccion
                                                        .Fields(7) = v12                                                            'udm
                                                        .Fields(8) = v3                                                             'folio
                                                        .Fields(9) = v18                                                            'cancelado
                                                        .Fields(10) = vLote                                                         'lote
                                                        .Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
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
                                                    .Fields(1) = v1                                     'idclienteproveedor
                                                    .Fields(2) = v2                                     'nombre cliente proveedor
                                                    .Fields(3) = v3                                     'folio
                                                    .Fields(4) = v4                                     'LugarVenta
                                                    .Fields(5) = v6                                     'fecha
                                                    .Fields(6) = v7                                     'Tipoarticulo
                                                    .Fields(7) = v8                                     'idarticulo
                                                    .Fields(8) = v9                                     'codigo articulo
                                                    .Fields(9) = v10                                    'descripcion articulo
                                                    .Fields(10) = CantidadRestante                      'cantidad
                                                    .Fields(11) = v12                                   'UDM
                                                    .Fields(12) = v13                                   'precio
                                                    .Fields(13) = v14                                   'subtotal
                                                    .Fields(14) = v15                                   'iva
                                                    .Fields(15) = v16                                   'total
                                                    .Fields(16) = v17                                   'totalpagado
                                                    .Fields(17) = v18                                   'cancelado
                                                    .Fields(18) = v19                                   'comentarios
                                                    .Fields(19) = v20                                   'tipo
                                                    .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                                    .Fields(21) = vLote                                 'lote
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
                                                        .Fields(1) = v8                                                     'id
                                                        .Fields(2) = v9                                                     'codigo
                                                        .Fields(3) = v10                                                    'descripcion
                                                        .Fields(4) = Date                                                   'fecha
                                                        .Fields(5) = "Devolución de venta"                                  'tipo de treansaccion
                                                        .Fields(7) = v12                                                    'udm
                                                        .Fields(8) = v3                                                     'folio
                                                        .Fields(9) = v18                                                    'cancelado
                                                        .Fields(10) = vLote                                                 'lote
                                                        .Fields(6) = Replace(Format(Val(vCantidadLote), "0.00"), ",", ".")  'cantidad
                                                    .Update
                                                    .Requery
                                                End With
                                            End If
                                            
                                            v14 = Replace(Format(Val(vCantidadLote) * Val(v13) * -1, "0.00"), ",", ".")             'subtotal
                                            v15 = Replace(Format(Val(vCantidadLote) * Val(v13) * Val(viva) * -1, "0.00"), ",", ".") 'iva
                                            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
                                            
                                            'guardar compra o venta
                                            With Rs3
                                                If .State = 1 Then .Close
                                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                                .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                                .Requery
                                                .AddNew
                                                    .Fields(1) = v1                                                         'idclienteproveedor
                                                    .Fields(2) = v2                                                         'nombre cliente proveedor
                                                    .Fields(3) = v3                                                         'folio
                                                    .Fields(4) = v4                                                         'LugarVenta
                                                    .Fields(5) = v6                                                         'fecha
                                                    .Fields(6) = v7                                                         'Tipoarticulo
                                                    .Fields(7) = v8                                                         'idarticulo
                                                    .Fields(8) = v9                                                         'codigo articulo
                                                    .Fields(9) = v10                                                        'descripcion articulo
                                                    .Fields(10) = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".") 'cantidad
                                                    .Fields(11) = v12                                                       'UDM
                                                    .Fields(12) = v13                                                       'precio
                                                    .Fields(13) = v14                                                       'subtotal
                                                    .Fields(14) = v15                                                       'iva
                                                    .Fields(15) = v16                                                       'total
                                                    .Fields(16) = v17                                                       'totalpagado
                                                    .Fields(17) = v18                                                       'cancelado
                                                    .Fields(18) = v19                                                       'comentarios
                                                    .Fields(19) = v20                                                       'tipo
                                                    .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "")                  'NUM_FOLIO
                                                    .Fields(21) = vLote                                                     'lote
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
                                                .Fields(1) = v8                                                 'id
                                                .Fields(2) = v9                                                 'codigo
                                                .Fields(3) = v10                                                'descripcion
                                                .Fields(4) = Date                                               'fecha
                                                .Fields(5) = "Devolución de venta"                              'tipo de treansaccion
                                                .Fields(6) = Replace(Format(Val(v11) * -1, "0.00"), ",", ".")   'cantidad
                                                .Fields(7) = v12                                                'udm
                                                .Fields(8) = v3                                                 'folio
                                                .Fields(9) = v18                                                'cancelado
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
                                            .Fields(1) = v1                                         'idclienteproveedor
                                            .Fields(2) = v2                                         'nombre cliente proveedor
                                            .Fields(3) = v3                                         'folio
                                            .Fields(4) = v4                                         'LugarVenta
                                            .Fields(5) = v6                                         'fecha
                                            .Fields(6) = v7                                         'Tipoarticulo
                                            .Fields(7) = v8                                         'idarticulo
                                            .Fields(8) = v9                                         'codigo articulo
                                            .Fields(9) = v10                                        'descripcion articulo
                                            .Fields(10) = v11                                       'cantidad
                                            .Fields(11) = v12                                       'UDM
                                            .Fields(12) = v13                                       'precio
                                            .Fields(13) = v14                                       'subtotal
                                            .Fields(14) = v15                                       'iva
                                            .Fields(15) = v16                                       'total
                                            .Fields(16) = v17                                       'totalpagado
                                            .Fields(17) = v18                                       'cancelado
                                            .Fields(18) = v19                                       'comentarios
                                            .Fields(19) = v20                                       'tipo
                                            .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "")  'NUM_FOLIO
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
                    v = 0 ' Ticket de pedido
                    
                    If Val(v11) > 0 Then
                        'si tiene control de lote
                        If ControlLote = 1 Then
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
                                                    .Fields(1) = v8                 'id
                                                    .Fields(2) = v9                 'codigo
                                                    .Fields(3) = v10                'descripcion
                                                    .Fields(4) = Date               'fecha
                                                    .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                                    .Fields(7) = v12                'udm
                                                    .Fields(8) = v3                 'folio
                                                    .Fields(9) = v18                'cancelado
                                                    .Fields(10) = vLote             'lote
                                                    .Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".") 'cantidad
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
                                                .Fields(1) = v1                 'idclienteproveedor
                                                .Fields(2) = v2                 'nombre cliente proveedor
                                                .Fields(3) = v3                 'folio
                                                .Fields(4) = v4                 'LugarVenta
                                                .Fields(5) = v6                 'fecha
                                                .Fields(6) = v7                 'Tipoarticulo
                                                .Fields(7) = v8                 'idarticulo
                                                .Fields(8) = v9                 'codigo articulo
                                                .Fields(9) = v10                'descripcion articulo
                                                .Fields(10) = CantidadRestante  'cantidad
                                                .Fields(11) = v12               'UDM
                                                .Fields(12) = v13               'precio
                                                .Fields(13) = v14               'subtotal
                                                .Fields(14) = v15               'iva
                                                .Fields(15) = v16               'total
                                                .Fields(16) = v17               'totalpagado
                                                .Fields(17) = v18               'cancelado
                                                .Fields(18) = v19               'comentarios
                                                .Fields(19) = v20               'tipo
                                                .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                                .Fields(21) = vLote             'lote
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
                                                    .Fields(1) = v8                 'id
                                                    .Fields(2) = v9                 'codigo
                                                    .Fields(3) = v10                'descripcion
                                                    .Fields(4) = Date               'fecha
                                                    .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                                    .Fields(7) = v12                'udm
                                                    .Fields(8) = v3                 'folio
                                                    .Fields(9) = v18                'cancelado
                                                    .Fields(10) = vLote             'lote
                                                    .Fields(6) = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".") 'cantidad
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
                                                .Fields(1) = v1             'idclienteproveedor
                                                .Fields(2) = v2             'nombre cliente proveedor
                                                .Fields(3) = v3             'folio
                                                .Fields(4) = v4             'LugarVenta
                                                .Fields(5) = v6             'fecha
                                                .Fields(6) = v7             'Tipoarticulo
                                                .Fields(7) = v8             'idarticulo
                                                .Fields(8) = v9             'codigo articulo
                                                .Fields(9) = v10            'descripcion articulo
                                                .Fields(10) = vCantidadLote 'cantidad
                                                .Fields(11) = v12           'UDM
                                                .Fields(12) = v13           'precio
                                                .Fields(13) = v14           'subtotal
                                                .Fields(14) = v15           'iva
                                                .Fields(15) = v16           'total
                                                .Fields(16) = v17           'totalpagado
                                                .Fields(17) = v18           'cancelado
                                                .Fields(18) = v19           'comentarios
                                                .Fields(19) = v20           'tipo
                                                .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                                .Fields(21) = vLote         'lote
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
                                        If Rs12.State = 1 Then Rs12.Close
                                        Rs12.CursorLocation = adodb.CursorLocationEnum.adUseClient
                                        Rs12.Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                        Rs12.Requery
                                        Rs12.AddNew
                                            Rs12.Fields(1) = v8 'idarticulo
                                            Rs12.Fields(2) = vCurrentLote 'lote
                                            Rs12.Fields(3) = "Venta" 'tipo"
                                        Rs12.Update
                                        Rs12.Requery
                                        Rs12.Close
                                    End If
                                    
                                    If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                        With Rs5
                                            If .State = 1 Then .Close
                                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                            .Requery
                                            .AddNew
                                                .Fields(1) = v8                 'id
                                                .Fields(2) = v9                 'codigo
                                                .Fields(3) = v10                'descripcion
                                                .Fields(4) = Date               'fecha
                                                .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                                .Fields(7) = v12                'udm
                                                .Fields(8) = v3                 'folio
                                                .Fields(9) = v18                'cancelado
                                                .Fields(10) = vCurrentLote      'lote
                                                .Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".") 'cantidad
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
                                            .Fields(1) = v1                 'idclienteproveedor
                                            .Fields(2) = v2                 'nombre cliente proveedor
                                            .Fields(3) = v3                 'folio
                                            .Fields(4) = v4                 'LugarVenta
                                            .Fields(5) = v6                 'fecha
                                            .Fields(6) = v7                 'Tipoarticulo
                                            .Fields(7) = v8                 'idarticulo
                                            .Fields(8) = v9                 'codigo articulo
                                            .Fields(9) = v10                'descripcion articulo
                                            .Fields(10) = CantidadRestante  'cantidad
                                            .Fields(11) = v12               'UDM
                                            .Fields(12) = v13               'precio
                                            .Fields(13) = v14               'subtotal
                                            .Fields(14) = v15               'iva
                                            .Fields(15) = v16               'total
                                            .Fields(16) = v17               'totalpagado
                                            .Fields(17) = v18               'cancelado
                                            .Fields(18) = v19               'comentarios
                                            .Fields(19) = v20               'tipo
                                            .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                            .Fields(21) = vCurrentLote      'lote
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
                            
                            If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                'salida por venta
                                With Rs5
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                        .Fields(1) = v8                 'id
                                        .Fields(2) = v9                 'codigo
                                        .Fields(3) = v10                'descripcion
                                        .Fields(4) = Date               'fecha
                                        .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                        .Fields(6) = Replace(Format(Val(v11) * -1, "0.00"), ",", ".") 'cantidad
                                        .Fields(7) = v12                'udm
                                        .Fields(8) = v3                 'folio
                                        .Fields(9) = v18                'cancelado
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
                                    .Fields(1) = v1     'idclienteproveedor
                                    .Fields(2) = v2     'nombre cliente proveedor
                                    .Fields(3) = v3     'folio
                                    .Fields(4) = v4     'LugarVenta
                                    .Fields(5) = v6     'fecha
                                    .Fields(6) = v7     'Tipoarticulo
                                    .Fields(7) = v8     'idarticulo
                                    .Fields(8) = v9     'codigo articulo
                                    .Fields(9) = v10    'descripcion articulo
                                    .Fields(10) = v11   'cantidad
                                    .Fields(11) = v12   'UDM
                                    .Fields(12) = v13   'precio
                                    .Fields(13) = v14   'subtotal
                                    .Fields(14) = v15   'iva
                                    .Fields(15) = v16   'total
                                    .Fields(16) = v17   'totalpagado
                                    .Fields(17) = v18   'cancelado
                                    .Fields(18) = v19   'comentarios
                                    .Fields(19) = v20   'tipo
                                    .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
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
                        Unload TicketPedido
                        
                        With TicketPedido
                            Set .DataSource = Rs6
                        
                            With .Sections("Sección1")
                                .Controls("Texto1").DataField = "cantidad"
                                .Controls("Texto2").DataField = "articulo"
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
        
        If Rs1.State = 1 Then Rs1.Close
        If Rs2.State = 1 Then Rs2.Close
        If Rs3.State = 1 Then Rs3.Close
        If Rs4.State = 1 Then Rs4.Close
        If Rs5.State = 1 Then Rs5.Close
        If Rs6.State = 1 Then Rs6.Close
        If Rs9.State = 1 Then Rs9.Close
        If Rs12.State = 1 Then Rs12.Close
        
        Set Rs1 = Nothing
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
