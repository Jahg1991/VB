VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHistorialCompras 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Historial de movimientos"
   ClientHeight    =   7215
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   15555
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
   ScaleWidth      =   15555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3240
      Index           =   0
      Left            =   4170
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
            Picture         =   "HistorialCompras.frx":0000
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
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   15015
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   0
            Left            =   1080
            TabIndex        =   14
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   151912449
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
            TabIndex        =   13
            Top             =   4200
            Width           =   14775
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
            TabIndex        =   12
            Top             =   1200
            Width           =   14775
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   4200
            TabIndex        =   16
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   151912449
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   3120
            Width           =   13095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"HistorialCompras.frx":080F
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
            TabIndex        =   18
            Top             =   3840
            Width           =   12855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"HistorialCompras.frx":089D
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
            TabIndex        =   17
            Top             =   840
            Width           =   14655
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
            TabIndex        =   15
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
Attribute VB_Name = "frmHistorialCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//OTROS
    Dim i                   As Long
    Dim Prt                 As Printer
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset  'Cabecera ComprasVentas
    Dim Rs1                 As New adodb.Recordset  'VentasCompras
    Dim Rs2                 As New adodb.Recordset  'Pagos
    Dim Rs3                 As New adodb.Recordset  'inventarios
    Dim Rs6                 As New adodb.Recordset  'ticket
    
    '//COMPRAS
    Dim c1                  As String
    Dim c2                  As String
    Dim c3                  As String
    Dim c4                  As String
    Dim c5                  As String
    Dim c6                  As String
    Dim c7                  As String
    Dim nc                  As Long
    
    '//OTROS
    Dim vTicketSubtotal     As String
    Dim vTicketIva          As String
    Dim vTicketTotal        As String
    
    '//PAGOS
    Dim DineroRestante      As String
    Dim TMovimiento         As String
    Dim TPagado             As String
    Dim TDebido             As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        For i = 0 To 1
            DTPicker1(i).Value = Date
        Next i
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            
            frmHistorialCompras.Pagar.Caption = "Pagar Compra"
                
            If StTipoCompra = "Pagadas" Then
                .Open "Select * from PO_HEADERS_ALL_P where totalPagado = total;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            End If
                
            If StTipoCompra = "No Pagadas" Then
                .Open "Select * from PO_HEADERS_ALL_P where totalPagado < total;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            End If
            
            .Requery
        End With
        
        If Rs.RecordCount > 0 Then
            Rs.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
            Rs.Requery
        
            List1.Clear
            List2.Clear
            
            Do Until Rs.EOF
                c1 = Mid(Rs!Folio, 1, 10)
                c2 = Mid(Rs!Fecha, 1, 10)
                c3 = Mid(Rs!Nombre, 1, 44)
                c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
                c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                c7 = Replace(Format(Mid(Rs!TotalDebido, 1, 12), "0.00"), ",", ".")
                    
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
                
                nc = 12 - Len(c7)
                
                For i = 1 To nc
                    c7 = " " & c7
                Next i
                
                List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                
                Rs.MoveNext
            Loop
            
            With Rs1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from PO_LINES_ALL", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            End With
            
            With Rs2
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from RA_CASH_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            End With
            
            With Rs3
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from MTL_MATERIAL_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            End With
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub DTPicker1_Change(Index As Integer)
        On Error GoTo errHandler
        
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
                    c7 = Replace(Format(Mid(Rs!TotalDebido, 1, 12), "0.00"), ",", ".")
                        
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
                    
                    nc = 12 - Len(c7)
                    
                    For i = 1 To nc
                        c7 = " " & c7
                    Next i
                    
                    List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                    
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
                    c7 = Replace(Format(Mid(Rs!TotalDebido, 1, 12), "0.00"), ",", ".")
                        
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
                    
                    nc = 12 - Len(c7)
                    
                    For i = 1 To nc
                        c7 = " " & c7
                    Next i
                    
                    List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                    
                    Rs.MoveNext
                Loop
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:DTPicker1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub List1_Click()
        On Error GoTo errHandler
        
        Label2.Caption = Get_Comentario(Trim(Mid(List1.Text, 1, 10)))
        List2.Clear
        
        Rs1.Requery
        Rs1.Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
        
        Do Until Rs1.EOF
            c1 = Mid(Rs1!DescripcionArticulo & " (" & Rs1!UDM & ")" & " (" & Rs1!CodigoArticulo & ")", 1, 27)
            c2 = Replace(Format(Mid(Rs1!cantidad, 1, 12), "0.00"), ",", ".")
            c3 = Replace(Format(Mid(Rs1!Precio, 1, 12), "0.00"), ",", ".")
            c4 = Replace(Format(Mid(Rs1!subtotal, 1, 12), "0.00"), ",", ".")
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
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:List1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Public Sub Ticket_Click()
        On Error GoTo errHandler
        
        If Mid(List1.Text, 1, 10) <> "" Then
            With Rs6
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from PO_TRANSACTION_TICKET where folio = '" & Trim(Mid(List1.Text, 1, 10)) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                    
                If .RecordCount <> 0 Then
                    With TicketComprasVentas
                        vTicketSubtotal = Get_SumSubtotal(Rs6.Fields(6))
                        vTicketIva = Get_SumIva(Rs6.Fields(6))
                        vTicketTotal = Get_SumTotal(Rs6.Fields(6))
                                
                        vTicketSubtotal = Replace(Format(Val(vTicketSubtotal), "0.00"), ",", ".")
                        vTicketIva = Replace(Format(Val(vTicketIva), "0.00"), ",", ".")
                        vTicketTotal = Replace(Format(Val(vTicketTotal), "0.00"), ",", ".")
                                
                        Set .DataSource = Rs6
                                
                        With .Sections("Sección4")
                            .Controls("Etiqueta2").Caption = "TICKET DE COMPRA"
                            .Controls("Etiqueta30").Caption = "Usuario: " & StUsuario
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
                            .Controls("Etiqueta23").Caption = "$ " & vTicketSubtotal    'subtotal
                            .Controls("Etiqueta26").Caption = "$ " & vTicketIva         'iva
                            .Controls("Etiqueta27").Caption = "$ " & vTicketTotal       'total
                            
                            .Controls("Label1").Visible = False
                            .Controls("Label2").Visible = False
                            .Controls("Label3").Visible = False
                            .Controls("Label4").Visible = False
                            .Controls("Label5").Visible = False
                            .Controls("Label6").Visible = False
                            .Controls("Etiqueta25").Visible = False
                            .Controls("Etiqueta28").Visible = False
                        End With
                                
                        .Show 1
                    End With
                End If
                        
                .Close
            End With
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Ticket_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Pagar_Click()
        On Error GoTo errHandler
        
        If Mid(List1.Text, 1, 10) <> "" Then
            Archivo.Enabled = False
            
            Frame1.Enabled = False
            
            Frame2(0).Visible = True
            
            With Rs1
                .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                .Requery
            End With
            
            Text2(0) = Replace(Format(Val(Mid(List1.Text, 68, 12)) - Val(Mid(List1.Text, 81, 12)), "0.00"), ",", ".")
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Pagar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Text2_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
                    
        If vbq = vbYes Then
            Frame1.Enabled = True
            Frame2(0).Visible = False
                
            If Val(Text2(1)) > 0 Then
                With Rs2
                    .AddNew
                        .Fields(1) = Date
                        .Fields(2) = "Pago de compra"
                            
                        If Val(Text2(1)) >= Val(Text2(0)) Then
                            .Fields(3) = Replace(Format(Val(Text2(0)) * -1, "0.00"), ",", ".")
                        Else
                            .Fields(3) = Replace(Format(Val(Text2(1)) * -1, "0.00"), ",", ".")
                        End If
                            
                        .Fields(4) = Trim(Mid(List1.Text, 1, 10))
                        .Fields(5) = "No"
                        .Fields(6) = frmMenuInicial.Combo1.Text
                    .Update
                    .Requery
                End With
                    
                With Rs1
                    DineroRestante = Val(Text2(1))
                        
                    Do Until .EOF
                        TMovimiento = Replace(.Fields(15), ",", ".")
                        TPagado = Replace(.Fields(16), ",", ".")
                        TDebido = Replace(Val(TMovimiento) - Val(TPagado), ",", ".")
                       
                        If Val(TDebido) > 0 Then
                            If DineroRestante > Val(TDebido) Then
                                .Fields(16) = .Fields(15)
                                DineroRestante = DineroRestante - Val(TDebido)
                            Else
                                .Fields(16) = Val(TPagado) + DineroRestante
                                DineroRestante = 0
                            End If
                        End If
                            
                        .MoveNext
                    Loop
                        
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
                    c4 = ""
                    c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
                    c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                    c7 = Replace(Format(Mid(Rs!TotalDebido, 1, 12), "0.00"), ",", ".")
                            
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
                        
                    nc = 12 - Len(c7)
                        
                    For i = 1 To nc
                        c7 = " " & c7
                    Next i
                        
                    List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                    
                    Rs.MoveNext
                Loop
                            
                Frame1.Enabled = True
                Frame2(0).Visible = False
            End If
        End If
            
        Archivo.Enabled = True
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Cancelar_Click()
        On Error GoTo errHandler
        
        If Mid(List1.Text, 1, 10) <> "" Then
            vbq = MsgBox("¿Desea cancelar el movimiento?", vbQuestion + vbYesNo, "Información")
                    
            If vbq = vbYes Then
                With Rs1
                    .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                    .Requery
                    
                    If Rs1.RecordCount > 0 Then
                        Do Until .EOF
                            .Fields(17) = "Si"
                            .MoveNext
                        Loop
                        
                        .Filter = ""
                        .Requery
                    End If
                End With
                
                With Rs2
                    .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                    .Requery
                    
                    If Rs2.RecordCount > 0 Then
                        Do Until .EOF
                            .Fields(5) = "Si"
                            .MoveNext
                        Loop
                    
                        .Filter = ""
                        .Requery
                    End If
                End With
                
                With Rs3
                    .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                    .Requery
                    
                    If Rs3.RecordCount > (0) Then
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
                    c3 = Mid(Rs!Nombre, 1, 44)
                    c5 = Replace(Format(Mid(Rs!Total, 1, 12), "0.00"), ",", ".")
                    c6 = Replace(Format(Mid(Rs!TotalPagado, 1, 12), "0.00"), ",", ".")
                    c7 = Replace(Format(Mid(Rs!TotalDebido, 1, 12), "0.00"), ",", ".")
                                
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
                            
                    nc = 12 - Len(c7)
                            
                    For i = 1 To nc
                        c7 = " " & c7
                    Next i
                            
                    List1.AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                    
                    Rs.MoveNext
                Loop
                
                Exit Sub
            Else
                Exit Sub
            End If
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Cancelar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Agregar_Click()
        On Error GoTo errHandler
        
        If Mid(List1.Text, 1, 10) <> "" Then
            With Rs1
                .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                .Requery
            End With
    
            With frmAgregarArticuloCompras
                .Caption = "Añadir artículos"
                
                .Text1(0) = Trim(Mid(List1.Text, 1, 10))
                .Text1(7) = Rs1.Fields(2).Value
                
                If IsNull(Rs1.Fields(4).Value) = False Then
                    .Text1(8) = Rs1.Fields(4).Value
                End If
                
                .Show 1
            End With
            
            Unload frmHistorialCompras
            
            Set frmHistorialCompras = Nothing
            
            frmHistorialCompras.Show
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Agregar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        Unload TicketComprasVentas
        
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
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialCompras:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
