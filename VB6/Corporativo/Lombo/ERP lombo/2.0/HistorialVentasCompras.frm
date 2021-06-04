VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHistorialVentasCompras 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Historial de movimientos"
   ClientHeight    =   7215
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13965
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
   ScaleHeight     =   7215
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2040
      Index           =   3
      Left            =   5055
      TabIndex        =   12
      Top             =   2587
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1815
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3615
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   4
            Left            =   1920
            Picture         =   "HistorialVentasCompras.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   240
            Picture         =   "HistorialVentasCompras.frx":068B
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "¿Desea cancelar el movimiento?"
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
            Height          =   975
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3240
      Index           =   0
      Left            =   3375
      TabIndex        =   3
      Top             =   1987
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3015
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6975
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   2760
            Picture         =   "HistorialVentasCompras.frx":0E9A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   2
            Left            =   1560
            TabIndex        =   7
            Top             =   1560
            Width           =   5055
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   0
            Left            =   1560
            TabIndex        =   5
            Top             =   120
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cambio"
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
            Index           =   13
            Left            =   -600
            TabIndex        =   11
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pagado"
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
            Index           =   12
            Left            =   -600
            TabIndex        =   10
            Top             =   960
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
            Index           =   11
            Left            =   -600
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13455
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   0
            Left            =   1080
            TabIndex        =   19
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   117440513
            CurrentDate     =   43915
         End
         Begin VB.ListBox List2 
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
            Height          =   2310
            Left            =   120
            TabIndex        =   18
            Top             =   4200
            Width           =   13095
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
            Height          =   1740
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   13095
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   4200
            TabIndex        =   21
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   117440513
            CurrentDate     =   43915
         End
         Begin VB.Label Label1 
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
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1800
            TabIndex        =   24
            Top             =   3120
            Width           =   11295
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"HistorialVentasCompras.frx":16A9
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
            Left            =   240
            TabIndex        =   23
            Top             =   3840
            Width           =   12855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"HistorialVentasCompras.frx":1737
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
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   12855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Left            =   3360
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Ticket 
         Caption         =   "Ticket"
         Shortcut        =   ^T
      End
      Begin VB.Menu Pagar 
         Caption         =   "Pagar"
         Shortcut        =   ^P
      End
      Begin VB.Menu Agregar 
         Caption         =   "Agregar Articulo"
         Shortcut        =   ^A
      End
      Begin VB.Menu Cancelar 
         Caption         =   "Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmHistorialVentasCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Rs As New ADODB.Recordset 'Cabecera ComprasVentas
Dim Rs1 As New ADODB.Recordset 'VentasCompras
Dim Rs2 As New ADODB.Recordset 'Pagos
Dim Rs3 As New ADODB.Recordset 'inventarios
Dim Rs6 As New ADODB.Recordset 'ticket
Dim c1 As String
Dim c2 As String
Dim c3 As String
Dim c4 As String
Dim c5 As String
Dim c6 As String
Dim nc As Integer

Private Sub Form_Load()

    On Error Resume Next
    
    For i = 0 To 1
        DTPicker1(i).Value = Date
    Next i
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs
        
        If .State = 1 Then .Close
        
        If StTipoVentasCompras = "Ventas" Then
        
            frmHistorialVentasCompras.Pagar.Caption = "Cobrar Venta"
        
            If StTipoVenta = "Local" Then
                .Open "Select * from CabeceraVentas where LugarVenta = 'Local';", Cn, adOpenStatic, adLockOptimistic
            End If
            
            If StTipoVenta = "Domicilio" Then
                .Open "Select * from CabeceraVentas where LugarVenta = 'Domicilio';", Cn, adOpenStatic, adLockOptimistic
            End If
            
            If StTipoVenta = "Abiertas" Then
                .Open "Select * from CabeceraVentas1 where totalPagado = 0;", Cn, adOpenStatic, adLockOptimistic
            End If
            
        End If
        
        If StTipoVentasCompras = "Compras" Then
        
            frmHistorialVentasCompras.Pagar.Caption = "Pagar Compra"
            
            If StTipoCompra = "Pagadas" Then
                .Open "Select * from CabeceraCompras where totalPagado > 0;", Cn, adOpenStatic, adLockOptimistic
            End If
            
            If StTipoCompra = "No Pagadas" Then
                .Open "Select * from CabeceraCompras where totalPagado = 0;", Cn, adOpenStatic, adLockOptimistic
            End If
            
        End If
        
        .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
        .Requery
        
    End With
    
    If Rs.RecordCount > 0 Then
    
        List1.Clear
        List2.Clear
        
        Do Until Rs.EOF
            c1 = Mid(Rs!Folio, 1, 10)
            c2 = Mid(Rs!Fecha, 1, 10)
            c3 = Mid(Rs!Nombre, 1, 44)
            c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
            c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                
            nc = 10 - Len(c1)
            For i = 1 To nc
                c1 = c1 & " "
            Next i
               
            nc = 10 - Len(c2)
            For i = 1 To nc
                c2 = c2 & " "
            Next i
            
            nc = 44 - Len(c3)
            For i = 1 To nc
                c3 = c3 & " "
            Next i
            
            nc = 12 - Len(c5)
            For i = 1 To nc
                c5 = " " & c5
            Next i
            
            nc = 12 - Len(c6)
            For i = 1 To nc
                c6 = " " & c6
            Next i
            
            List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6
                
            Rs.MoveNext
            
        Loop
        
        With Rs1
            
            If .State = 1 Then .Close
            .Open "select * from HistorialVentasCompras", Cn, adOpenStatic, adLockOptimistic
            .Requery
            
        End With
        
        With Rs2
                    
            If .State = 1 Then .Close
            .Open "Select * from MovimientosCaja", Cn, adOpenStatic, adLockOptimistic
            .Requery
        
        End With
        
        With Rs3
                    
            If .State = 1 Then .Close
            .Open "Select * from TransaccionesDeInventario", Cn, adOpenStatic, adLockOptimistic
            .Requery
        
        End With
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Información"
        frmMenuInicial.Enabled = True
        Unload Me
        
    End If
    
End Sub

Private Sub DTPicker1_Change(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            
            With Rs
        
                .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                .Requery
                
            End With
            
            List1.Clear
            List2.Clear
            
            Do Until Rs.EOF
                c1 = Mid(Rs!Folio, 1, 10)
                c2 = Mid(Rs!Fecha, 1, 10)
                c3 = Mid(Rs!Nombre, 1, 44)
                c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
                c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                    
                nc = 10 - Len(c1)
                For i = 1 To nc
                    c1 = c1 & " "
                Next i
                   
                nc = 10 - Len(c2)
                For i = 1 To nc
                    c2 = c2 & " "
                Next i
                
                nc = 44 - Len(c3)
                For i = 1 To nc
                    c3 = c3 & " "
                Next i
                
                nc = 12 - Len(c5)
                For i = 1 To nc
                    c5 = " " & c5
                Next i
                
                nc = 12 - Len(c6)
                For i = 1 To nc
                    c6 = " " & c6
                Next i
                
                List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6
                    
                Rs.MoveNext
                
            Loop
            
        Case 1
            
            With Rs
        
                .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                .Requery
                
            End With
            
            List1.Clear
            List2.Clear
            
            Do Until Rs.EOF
                c1 = Mid(Rs!Folio, 1, 10)
                c2 = Mid(Rs!Fecha, 1, 10)
                c3 = Mid(Rs!Nombre, 1, 44)
                c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
                c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                    
                nc = 10 - Len(c1)
                For i = 1 To nc
                    c1 = c1 & " "
                Next i
                   
                nc = 10 - Len(c2)
                For i = 1 To nc
                    c2 = c2 & " "
                Next i
                
                nc = 44 - Len(c3)
                For i = 1 To nc
                    c3 = c3 & " "
                Next i
                    
                If StTipoVentasCompras = "Ventas" Then
                    nc = 10 - Len(c4)
                    For i = 1 To nc
                        c4 = c4 & " "
                    Next i
                End If
                
                nc = 12 - Len(c5)
                For i = 1 To nc
                    c5 = " " & c5
                Next i
                
                nc = 12 - Len(c6)
                For i = 1 To nc
                    c6 = " " & c6
                Next i
                
                List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6
                    
                Rs.MoveNext
                
            Loop
    
    End Select
    
End Sub

Private Sub List1_Click()

    On Error Resume Next
    
    Label2.Caption = Get_Comentario(Trim(Mid(List1.Text, 1, 10)))
    
    List2.Clear
    
    Rs1.Requery
    
    Rs1.Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
            
    Do Until Rs1.EOF
    
        c1 = Mid(Rs1!DescripcionArticulo & " (" & Rs1!UDM & ")", 1, 27)
        c2 = Replace(Format(Mid(Rs1!Cantidad, 1, 12), "0.00"), ",", ".")
        c3 = Replace(Format(Mid(Rs1!Precio, 1, 12), "0.00"), ",", ".")
        c4 = Replace(Format(Mid(Rs1!Subtotal, 1, 12), "0.00"), ",", ".")
        c5 = Replace(Format(Mid(Rs1!IVA, 1, 12), "0.00"), ",", ".")
        c6 = Replace(Format(Mid(Rs1!Total, 1, 12), "0.00"), ",", ".")
                        
        nc = 27 - Len(c1)
        For i = 1 To nc
            c1 = c1 & " "
        Next i
        
        nc = 12 - Len(c2)
        For i = 1 To nc
            c2 = " " & c2
        Next i
        
        nc = 12 - Len(c3)
        For i = 1 To nc
            c3 = " " & c3
        Next i
        
        nc = 12 - Len(c4)
        For i = 1 To nc
            c4 = " " & c4
        Next i
        
        nc = 12 - Len(c5)
        For i = 1 To nc
            c5 = " " & c5
        Next i
        
        nc = 12 - Len(c6)
        For i = 1 To nc
            c6 = " " & c6
        Next i
                        
        List2.AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5 & " " & c6
        
        Rs1.MoveNext
                
    Loop
                
End Sub

Public Sub Ticket_Click()

    On Error Resume Next
    
    Dim vTicketSubtotal As String
    Dim vTicketIva As String
    Dim vTicketTotal As String
    Dim Prt As Printer
    
    If Mid(List1.Text, 1, 10) <> "" Then
        
        With Rs6
            
            If .State = 1 Then .Close
            .Open "Select * from ticket where folio = '" & Mid(List1.Text, 1, 10) & "';", Cn, adOpenStatic, adLockOptimistic
            .Requery
                
            If .RecordCount <> 0 Then
                
                Unload TicketComprasVentas
                
                With TicketComprasVentas
                        
                    vTicketSubtotal = Get_SumSubtotal(Rs6.Fields(6))
                    vTicketIva = Get_SumIva(Rs6.Fields(6))
                    vTicketTotal = Get_SumTotal(Rs6.Fields(6))
                        
                    vTicketSubtotal = Replace(Format(Val(vTicketSubtotal), "0.00"), ",", ".")
                    vTicketIva = Replace(Format(Val(vTicketIva), "0.00"), ",", ".")
                    vTicketTotal = Replace(Format(Val(vTicketTotal), "0.00"), ",", ".")
                        
                    Set .DataSource = Rs6
                        
                    With .Sections("Sección4")
                        
                        If StTipoVentasCompras = "Ventas" Then
                            .Controls("Etiqueta2").Caption = "TICKET DE VENTA"
                        Else
                            .Controls("Etiqueta2").Caption = "TICKET DE COMPRA"
                        End If
                        
                        .Controls("Etiqueta3").Caption = PcNombreEmpresa
                        .Controls("Etiqueta4").Caption = PcRFC
                        .Controls("Etiqueta5").Caption = PcDireccion
                        .Controls("Etiqueta6").Caption = PcTelefono
                        .Controls("Etiqueta11").Caption = Rs6.Fields(2) 'cliente
                        .Controls("Etiqueta12").Caption = Rs6.Fields(3) 'calle
                        .Controls("Etiqueta13").Caption = Rs6.Fields(4) 'colonia
                        .Controls("Etiqueta14").Caption = Rs6.Fields(5) 'telefono
                        .Controls("Etiqueta17").Caption = Rs6.Fields(7) 'fecha
                        .Controls("Etiqueta18").Caption = Rs6.Fields(6) 'folio
                    End With
                        
                    With .Sections("Sección1")
                        .Controls("Texto1").DataField = "cantidad"
                        .Controls("Texto2").DataField = "articulo"
                        .Controls("Texto3").DataField = "subtotal"
                    End With
                            
                    With .Sections("Sección5")
                        .Controls("Etiqueta23").Caption = "$ " & vTicketSubtotal 'subtotal
                        .Controls("Etiqueta26").Caption = "$ " & vTicketIva 'iva
                        .Controls("Etiqueta27").Caption = "$ " & vTicketTotal 'total
                    End With
                        
                    .Hide
                        
                    .PrintReport True
                        
                End With
                    
                Unload TicketComprasVentas
                    
            End If
                
            .Close
                
        End With
    
    Else
        
        MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
    
    End If
        
End Sub

Private Sub Pagar_Click()

    On Error Resume Next
    
    If Mid(List1.Text, 1, 10) <> "" Then
    
        Frame1.Enabled = False
        Frame2(0).Visible = True
        
        With Rs1
            .Filter = "Folio = '" & Mid(List1.Text, 1, 10) & "'"
            .Requery
        End With
        
        Text2(0) = Mid(List1.Text, 68, 12)
        Text2(1) = 0
    
    Else
        
        MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
    
    End If
    
End Sub

Private Sub Text2_Change(Index As Integer)

    On Error Resume Next
    
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
            
End Sub

Private Sub Command2_Click()

    On Error Resume Next
    
    If Val(Text2(1)) = 0 Or Val(Text2(1)) >= Val(Text2(0)) Then
        
        Frame1.Enabled = True
        Frame2(0).Visible = False
        
        If Val(Text2(1)) >= Val(Text2(0)) Then
            
            With Rs2
                
                .AddNew
                    .Fields(1) = Date
                    
                    If StTipoVentasCompras = "Ventas" Then
                        .Fields(2) = "Pago de venta"
                        .Fields(3) = Replace(Format(Val(Text2(0)), "0.00"), ",", ".")
                    End If
                    
                    If StTipoVentasCompras = "Compras" Then
                        .Fields(2) = "Pago de compra"
                        .Fields(3) = Replace(Format(Val(Text2(0)) * -1, "0.00"), ",", ".")
                    End If
                    
                    .Fields(4) = Mid(List1.Text, 1, 10)
                    .Fields(5) = "No"
                    
                .Update
                .Requery
                
            End With
            
            With Rs1
            
                Do Until .EOF
                    .Fields(17) = .Fields(16)
                    .MoveNext
                Loop
            
                .Update
                .Requery
            
            End With
            
            With Rs
            
                .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                .Requery
                
            End With
    
            List1.Clear
            List2.Clear
            
            Do Until Rs.EOF
                c1 = Mid(Rs!Folio, 1, 10)
                c2 = Mid(Rs!Fecha, 1, 10)
                c3 = Mid(Rs!Nombre, 1, 44)
                c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
                c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                    
                nc = 10 - Len(c1)
                For i = 1 To nc
                    c1 = c1 & " "
                Next i
                   
                nc = 10 - Len(c2)
                For i = 1 To nc
                    c2 = c2 & " "
                Next i
                
                nc = 44 - Len(c3)
                For i = 1 To nc
                    c3 = c3 & " "
                Next i
                
                nc = 12 - Len(c5)
                For i = 1 To nc
                    c5 = " " & c5
                Next i
                
                nc = 12 - Len(c6)
                For i = 1 To nc
                    c6 = " " & c6
                Next i
                
                List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6
                    
                Rs.MoveNext
                
            Loop
                    
            Frame1.Enabled = True
            Frame2(0).Visible = False
            
        End If
        
    End If
        
End Sub

Private Sub Cancelar_Click()

    On Error Resume Next
    
    If Mid(List1.Text, 1, 10) <> "" Then
        Frame1.Enabled = False
        Frame2(3).Visible = True
    Else
        MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
    End If
    
End Sub

Private Sub Command3_Click()

    On Error Resume Next
    
    With Rs1
        .Filter = "Folio = '" & Mid(List1.Text, 1, 10) & "'"
        .Requery
        
        If Rs1.RecordCount > 0 Then
            Do Until .EOF
                .Fields(18) = "Si"
                .MoveNext
            Loop
            
            '.Update
            .Filter = ""
            .Requery
        End If
        
    End With
    
    With Rs2
        .Filter = "Folio = '" & Mid(List1.Text, 1, 10) & "'"
        .Requery
        
        If Rs2.RecordCount > 0 Then
            Do Until .EOF
                .Fields(5) = "Si"
                .MoveNext
            Loop
        
            '.Update
            .Filter = ""
            .Requery
        End If
        
    End With
    
    With Rs3
        .Filter = "Folio = '" & Mid(List1.Text, 1, 10) & "'"
        .Requery
        
        If Rs3.RecordCount > 0 Then
            Do Until .EOF
                .Fields(9) = "Si"
                .MoveNext
            Loop

            .Filter = ""
            .Requery
        End If
        
    End With
    
    MsgBox "Movimiento Cancelado", vbOKOnly, "Terminado"
    
    With Rs
            
        .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
        .Requery
                
    End With
    
    List1.Clear
    List2.Clear
            
    Do Until Rs.EOF
        c1 = Mid(Rs!Folio, 1, 10)
        c2 = Mid(Rs!Fecha, 1, 10)
                
        If StTipoVentasCompras = "Ventas" Then
            c3 = Mid(Rs!Nombre, 1, 33)
            c4 = Mid(Rs!Mesa, 1, 10)
        End If
                
        If StTipoVentasCompras = "Compras" Then
            c3 = Mid(Rs!Nombre, 1, 44)
        End If
                
        c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
        c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                    
        nc = 10 - Len(c1)
        For i = 1 To nc
            c1 = c1 & " "
        Next i
                   
        nc = 10 - Len(c2)
        For i = 1 To nc
            c2 = c2 & " "
        Next i
                
        nc = 44 - Len(c3)
        For i = 1 To nc
            c3 = c3 & " "
        Next i
                
        nc = 12 - Len(c5)
        For i = 1 To nc
            c5 = " " & c5
        Next i
                
        nc = 12 - Len(c6)
        For i = 1 To nc
            c6 = " " & c6
        Next i
                
        List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6
                    
        Rs.MoveNext
                
    Loop
                    
    Frame1.Enabled = True
    Frame2(3).Visible = False

End Sub

Private Sub Command1_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
    
        Case 4
        
            Frame1.Enabled = True
            Frame2(3).Visible = False
    
    End Select
    
End Sub

Private Sub Agregar_Click()
    
    On Error Resume Next
    
    If Mid(List1.Text, 1, 10) <> "" Then
    
        With Rs1
            .Filter = "Folio = '" & Mid(List1.Text, 1, 10) & "'"
            .Requery
        End With

        frmAgregarArticuloVentasCompras.Caption = "Añadir artículos"
        frmAgregarArticuloVentasCompras.Text1(0) = Mid(List1.Text, 1, 10)
        frmAgregarArticuloVentasCompras.Text1(7) = Rs1.Fields(2).Value
        frmAgregarArticuloVentasCompras.Text1(8) = Rs1.Fields(4).Value
        frmAgregarArticuloVentasCompras.Text1(9) = Rs1.Fields(5).Value
        frmAgregarArticuloVentasCompras.Show 1
        
        Unload frmHistorialVentasCompras
        Set frmHistorialVentasCompras = Nothing
        frmHistorialVentasCompras.Show
    Else
        
        MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
    
    End If
    
End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    frmMenuInicial.Enabled = True
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Rs1.State = 1 Then Rs1.Close
    If Rs2.State = 1 Then Rs2.Close
    If Rs3.State = 1 Then Rs3.Close
    If Rs6.State = 1 Then Rs6.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Rs1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Rs6 = Nothing
    Set Cn = Nothing
    
End Sub
