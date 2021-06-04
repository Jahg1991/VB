VERSION 5.00
Begin VB.Form frmAgregarArticuloVentasCompras 
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
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7815
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   13335
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   9
            Left            =   6000
            TabIndex        =   7
            Top             =   2520
            Width           =   3015
         End
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
            TabIndex        =   26
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   10
            Top             =   3960
            Width           =   12855
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   1800
            Picture         =   "AgregarArticuloVentasCompras.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "AgregarArticuloVentasCompras.frx":08D1
            Style           =   1  'Graphical
            TabIndex        =   8
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
            Caption         =   $"AgregarArticuloVentasCompras.frx":1105
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
            TabIndex        =   27
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   7080
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mesa"
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
            Index           =   8
            Left            =   3840
            TabIndex        =   19
            Top             =   2520
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
Attribute VB_Name = "frmAgregarArticuloVentasCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs1 As New ADODB.Recordset 'clientesproveedores
Dim Rs2 As New ADODB.Recordset 'items
Dim Rs3 As New ADODB.Recordset 'ventascompras
Dim Rs4 As New ADODB.Recordset 'lista de ingredientes
Dim Rs5 As New ADODB.Recordset 'movimientos de inventarios
Dim Rs6 As New ADODB.Recordset 'pedido barra
Dim Rs7 As New ADODB.Recordset 'pedido cocina
Dim Rs9 As New ADODB.Recordset 'entrada de dinero
Dim IdTransaccion As Integer 'folio
Dim v1 As Integer 'idclienteproveedor
Dim v2 As String 'nombre cliente proveedor
Dim v3 As String 'folio
Dim v4 As String 'LugarVenta
Dim v5 As String 'mesa
Dim v6 As Date 'fecha
Dim v7 As String 'Tipoarticulo
Dim v8 As Integer 'idarticulo
Dim v9 As String 'codigo articulo
Dim v10 As String 'descripcion articulo
Dim v11 As String 'cantidad
Dim v12 As String 'UDM
Dim v13 As String 'precio
Dim v14 As String 'subtotal
Dim v15 As String 'iva
Dim v16 As String 'total
Dim v17 As String 'totalpagado
Dim v18 As String 'cancelado
Dim v19 As String 'comentarios
Dim v20 As String 'tipo
Dim InItemId As Integer
' Constantes para indicar el color de fondo del combobox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const COLOR_NO_ENCONTRADO = &HC0C0FF ' color cuando no se encontró
Const COLOR_NORMAL = &HE0E0E0 ' color cuando hay coincidencia


Private Sub Form_Load()

    On Error Resume Next
    
    Dim i As Integer
    
    For i = 1 To 3
        Text1(i).BackColor = &HC0C0FF
    Next i
    
    Combo1(1).BackColor = &HC0C0FF
    
    If StTipoVentasCompras = "Ventas" Then
        
        Label1(6).Visible = True
        Label1(8).Visible = True
        Text1(8).Visible = True
        Text1(9).Visible = True
    
    End If
    
    If StTipoVentasCompras = "Compras" Then
        Label1(6).Visible = False
        Label1(8).Visible = False
        Text1(8).Visible = False
        Text1(9).Visible = False
    End If
    
    With Rs1
        
        If .State = 1 Then .Close
            
            If StTipoVentasCompras = "Ventas" Then
                .Open "Select * from ClientesProveedores where tipo = 'Cliente' order by 2;", Cn, adOpenStatic, adLockOptimistic
            End If
            
            If StTipoVentasCompras = "Compras" Then
                .Open "Select * from ClientesProveedores where tipo = 'Proveedor' order by 2;", Cn, adOpenStatic, adLockOptimistic
            End If
        
        .Filter = ""
        .Requery
    
    End With
    
    With Rs2
        
        If .State = 1 Then .Close
        
        .Open "Select * from items order by 3;", Cn, adOpenStatic, adLockOptimistic
        .Filter = ""
        .Requery
        
        If .RecordCount <> 0 Then
            
            Combo1(1).Clear
            
            While Not .EOF
                Combo1(1).AddItem .Fields(2) & " (" & .Fields(5) & ")"
                .MoveNext
            Wend
        
        End If
    
    End With

End Sub

Private Sub Combo1_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 1
            
            With Combo1(1)
                
                If .Text = "" Then
                    
                    .BackColor = &HC0C0FF
                    
                    With Rs2
                        .Filter = ""
                        .Requery
                    End With
                    
                    With Text1(2)
                        .Text = ""
                    End With
                
                Else
                
                    .BackColor = &HE0E0E0
                    
                    InItemId = Get_ItemId(.Text)
                    
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With
                    
                    With Text1(2)
                        .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                    End With
                
                End If
            
            End With
    
    End Select

End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    Static cadena As String
    Dim i As Long
    
    Select Case Index
        
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
                
                Else
                    
                    ' Backcolor normal cuando hay coincidencia
                    .BackColor = COLOR_NORMAL
                    
                    InItemId = Get_ItemId(.Text)
                    
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With
                    
                    With Text1(2)
                        .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                    End With
                
                End If
            
            End With
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)
    
    On Error Resume Next
    
    Dim i As Integer
    Dim X As Integer
    Dim intX As Integer
    
    Dim listSubtotal As String
    Dim listIva As String
    Dim listTotal As String
    Dim vLstCantidad As String
    Dim vLstPrecio As String
    Dim viva As String
    
    Select Case Index
    
        Case 0
                
             If Combo1(1) <> "" And Text1(1) <> "" And Text1(2) <> "" Then
                    
                Dim viid As String
                Dim videscripcion As String
                Dim vicantidad As String
                Dim viprecio As String
                    
                Dim c1 As Integer
                Dim c2 As Integer
                Dim c3 As Integer
                Dim c4 As Integer
                    
                viid = Rs2.Fields(0).Value
                videscripcion = Mid(Combo1(1), 1, 47)
                vicantidad = Replace(Format(Val(Text1(1)), "0.00"), ",", ".")
                viprecio = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                    
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
                    .BackColor = &HC0C0FF
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
                    listIva = (((Val(vLstCantidad) * Val(vLstPrecio))) * Val(viva)) + listIva
                        
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
                .RemoveItem (intX)
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
                listIva = (((Val(vLstCantidad) * Val(vLstPrecio))) * Val(viva)) + listIva
                        
            Next i
                        
            listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                    
            listIva = Replace(Format(listIva, "0.00"), ",", ".")
            listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
            listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                    
            Text1(6) = listSubtotal
            Text1(5) = listIva
            Text1(4) = listTotal
    
    End Select

End Sub

Private Sub Text1_Change(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 1
            
            With Text1(1)
                
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            
            End With
        
        Case 2
            
            With Text1(2)
                
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            
            End With
            
        Case 3
            
            With Text1(3)
                
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            
            End With
            
        Case 7
        
            With Text1(7)
                
                If .Text = "" Then
                    
                    .BackColor = &HC0C0FF
                    
                    With Rs1
                        
                        .Filter = ""
                        .Requery
                        
                        v1 = 0
                        v2 = ""
                    
                    End With
                
                Else
                    
                    .BackColor = &HE0E0E0
                    
                    With Rs1
                        
                        .Filter = ""
                        .Filter = "nombre like '" & Text1(7) & "'"
                        .Requery
                        
                        v1 = .Fields(0)
                        v2 = .Fields(1)
                    
                    End With
                
                End If
            
            End With
    
    End Select

End Sub

Private Sub Guardar_Click()

    'On Error Resume Next

    Dim i As Integer
    Dim nLineaBarra As Integer
    Dim nLineaCocina As Integer
    
    'Mostrar pago
    If List1.ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
        
        'asignar valores a campos cabecera
        v3 = Text1(0) 'folio
        v6 = Date 'fecha
        v18 = "No" 'cancelado
        v19 = Text1(3) 'comentarios
        v20 = StTipoVentasCompras 'tipo
        v4 = Text1(8) 'LugarVenta
        If Text1(8) = "Local" Then
            v5 = Text1(9) 'mesa
        End If
        
        nLineaCocina = 0
        nLineaBarra = 0
        
        For i = 0 To List1.ListCount - 1
                
            List1.ListIndex = i
        
            'asignar valores a campos lineas
            v8 = Trim(Mid(List1.Text, 1, 10)) 'idarticulo
            v7 = Get_ItemTipo(v8) 'Tipoarticulo
            
            If v7 = "Barra" Then
                v7 = "Barra"
            Else
                v7 = "Cocina"
            End If
            
            v9 = Get_ItemCod(v8) 'codigo articulo
            v10 = Get_ItemDesc(v8) 'descripcion articulo
            v11 = Replace(Format(Val(Trim(Mid(List1.Text, 60, 15))), "0.00"), ",", ".") 'cantidad
            v12 = Get_ItemUDM(v8) 'UDM
            v13 = Replace(Format(Val(Trim(Mid(List1.Text, 76, 15))), "0.00"), ",", ".") 'precio
            v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")   'subtotal
            
            'AQUI ME QUEDE
            
            Dim viva As String
            
            viva = Get_ItemIva(v8)
            
            v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".") 'total
            v17 = 0 'totalpagado
        
            With Rs5
                If .State = 1 Then .Close
                .Open "Select * from TransaccionesDeInventario;", Cn, adOpenStatic, adLockOptimistic
                .Requery
            End With
            
            If StTipoVentasCompras = "Ventas" Then
            
                'lista de materiales
                Dim InLmExists As Integer
                
                With Rs4
                        
                    If .State = 1 Then .Close
                    .Open "Select * from ListasDeIngredientes where cantidad <> 0 and ItemPTId = " & v8 & ";", Cn, adOpenStatic, adLockOptimistic
                    .Requery
                        
                    If .RecordCount <> 0 Then
                            
                        While Not .EOF
                            Rs5.AddNew
                                Rs5.Fields(1) = .Fields(3).Value 'id
                                Rs5.Fields(2) = .Fields(4).Value 'codigo
                                Rs5.Fields(3) = .Fields(5).Value 'descripcion
                                Rs5.Fields(4) = Date 'fecha
                                Rs5.Fields(5) = "Consumo de Ingredientes" 'tipo de treansaccion
                                Rs5.Fields(6) = Replace(Format(.Fields(6).Value * Val(v11) * -1, "0.00"), ",", ".") 'cantidad
                                Rs5.Fields(7) = Get_ItemUDM(.Fields(3).Value) 'udm
                                Rs5.Fields(8) = v3 'folio
                                Rs5.Fields(9) = v18 'cancelado
                            Rs5.Update
                            Rs5.Requery
                                    
                            .MoveNext
                        Wend
                
                    Else
                        'salida por venta
                        Rs5.AddNew
                            Rs5.Fields(1) = v8 'id
                            Rs5.Fields(2) = v9 'codigo
                            Rs5.Fields(3) = v10 'descripcion
                            Rs5.Fields(4) = Date 'fecha
                            Rs5.Fields(5) = "Salida por venta" 'tipo de treansaccion
                            Rs5.Fields(6) = Replace(Format(Val(v11) * -1, "0.00"), ",", ".") 'cantidad
                            Rs5.Fields(7) = v12 'udm
                            Rs5.Fields(8) = v3 'folio
                            Rs5.Fields(9) = v18 'cancelado
                        Rs5.Update
                        Rs5.Requery
                    End If
                    
                End With
                
            End If
            
            If StTipoVentasCompras = "Compras" Then
                With Rs5
                        .AddNew
                        .Fields(1) = v8 'id
                        .Fields(2) = v9 'codigo
                        .Fields(3) = v10 'descripcion
                        .Fields(4) = Date 'fecha
                        .Fields(5) = "Recepción de compra" 'tipo de treansaccion
                        .Fields(6) = Replace(Format(Val(v11), "0.00"), ",", ".") 'cantidad
                        .Fields(7) = v12 'udm
                        .Fields(8) = v3 'folio
                        .Fields(9) = v18 'cancelado
                    .Update
                    .Requery
                End With
            End If
            
            Rs5.Close
            
            'guardar compra o venta
            With Rs3
            
                If .State = 1 Then .Close
                .Open "Select * from HistorialVentasCompras;", Cn, adOpenStatic, adLockOptimistic
                .Requery
                
                .AddNew
                    .Fields(1) = v1 'idclienteproveedor
                    .Fields(2) = v2 'nombre cliente proveedor
                    .Fields(3) = v3 'folio
        
                    If StTipoVentasCompras = "Ventas" Then
                        .Fields(4) = v4 'LugarVenta
                        
                        If v4 = "Local" Then
                            .Fields(5) = v5 'mesa
                        End If
                    End If
                    
                    .Fields(6) = v6 'fecha
                    .Fields(7) = v7 'Tipoarticulo
                    .Fields(8) = v8 'idarticulo
                    .Fields(9) = v9 'codigo articulo
                    .Fields(10) = v10 'descripcion articulo
                    .Fields(11) = v11 'cantidad
                    .Fields(12) = v12 'UDM
                    .Fields(13) = v13 'precio
                    .Fields(14) = v14 'subtotal
                    .Fields(15) = v15 'iva
                    .Fields(16) = v16 'total
                    .Fields(17) = v17 'totalpagado
                    .Fields(18) = v18 'cancelado
                    .Fields(19) = v19 'comentarios
                    .Fields(20) = v20 'tipo
                    .Fields(21) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                .Update
                .Requery
                .Close
        
            End With
            
            If v7 = "Barra" Then
                nLineaBarra = nLineaBarra + 1
            End If
            
            If v7 = "Cocina" Then
                nLineaCocina = nLineaCocina + 1
            End If
        
        Next i
        
        If StTipoVentasCompras = "Ventas" Then
        
            If nLineaBarra <> 0 Then
                'barra
                With Rs6
                    
                    If .State = 1 Then .Close
                    .Open "Select top " & nLineaBarra & " * from HistorialVentasCompras where tipoarticulo ='Barra' and folio = '" & Text1(0) & "' order by id desc;", Cn, adOpenStatic, adLockOptimistic
                    .Requery
                        
                    If .RecordCount <> 0 Then
                    
                        Unload TicketPedidos
                        
                        With TicketPedidos
    
                            Set .DataSource = Rs6
                                
                            With .Sections("Sección4")
                                .Controls("Etiqueta18").Caption = Rs6.Fields(3) 'folio
                            End With
                            
                            With .Sections("Sección2")
                                .Controls("Label3").Caption = Text1(9).Text  'mesa
                                .Controls("Label4").Caption = Text1(3).Text 'comentarios
                            End With
                                
                            With .Sections("Sección1")
                                .Controls("Texto1").DataField = "cantidad"
                                .Controls("Texto2").DataField = "DescripcionArticulo"
                            End With
                                    
                            Establecer (PcImpresoraCocina) 'pedido
                            
                            If EncontroImpresora = 0 Then
                                .Hide
                                .PrintReport
                            Else
                                .Show 1
                            End If
                                
                        End With
                            
                    End If
                        
                    .Close
                        
                End With
            End If
                
            If nLineaCocina <> 0 Then
                'cocina
                With Rs7
                    
                    If .State = 1 Then .Close
                    .Open "Select top " & nLineaCocina & " * from HistorialVentasCompras where tipoarticulo <>'Barra' and folio = '" & Text1(0) & "' order by id desc;", Cn, adOpenStatic, adLockOptimistic
                    .Requery
                        
                    If .RecordCount <> 0 Then
                    
                        Unload TicketPedidos
                        
                        With TicketPedidos
                            
                            Set .DataSource = Rs7
                                
                            With .Sections("Sección4")
                                .Controls("Etiqueta18").Caption = Rs7.Fields(3) 'folio
                            End With
                            
                            With .Sections("Sección2")
                                .Controls("Label3").Caption = Text1(9).Text   'mesa
                                .Controls("Label4").Caption = Text1(3).Text 'comentarios
                            End With
                                
                            With .Sections("Sección1")
                                .Controls("Texto1").DataField = "cantidad"
                                .Controls("Texto2").DataField = "DescripcionArticulo"
                            End With
                                    
                            Establecer (PcImpresoraCocina) 'pedido
                            
                            If EncontroImpresora = 0 Then
                                .Hide
                                .PrintReport
                            Else
                                .Show 1
                            End If
                                
                        End With
                            
                    End If
                        
                    .Close
                        
                End With
            End If
        End If
    
        With frmHistorialVentasCompras
            .Enabled = True
        End With
        
        Unload Me
    
    Else
        
        MsgBox "Llenar todos los campos", vbCritical, "Advertencia"
        
        Exit Sub
        
    End If
    
End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs1.State = 1 Then Rs1.Close
    If Rs2.State = 1 Then Rs2.Close
    If Rs3.State = 1 Then Rs3.Close
    If Rs4.State = 1 Then Rs4.Close
    If Rs5.State = 1 Then Rs5.Close
    If Rs6.State = 1 Then Rs5.Close
    If Rs7.State = 1 Then Rs5.Close
    If Rs9.State = 1 Then Rs9.Close
    
    Set Rs1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Rs4 = Nothing
    Set Rs5 = Nothing
    Set Rs6 = Nothing
    Set Rs7 = Nothing
    Set Rs9 = Nothing
    
End Sub


