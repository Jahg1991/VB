VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHistorialVentas 
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
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5280
      Index           =   0
      Left            =   4170
      TabIndex        =   17
      Top             =   967
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5055
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   6975
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   2760
            Picture         =   "HistorialVentas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   4320
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
            Left            =   1680
            TabIndex        =   24
            Top             =   3720
            Width           =   4935
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
            Left            =   1680
            TabIndex        =   23
            Top             =   3000
            Width           =   4935
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
            Left            =   1680
            TabIndex        =   22
            Top             =   2280
            Width           =   4935
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
            Index           =   3
            Left            =   1680
            TabIndex        =   21
            Top             =   1560
            Width           =   4935
         End
         Begin VB.ComboBox Combo2 
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
            Height          =   480
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   4935
         End
         Begin VB.TextBox Text2 
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
            Index           =   4
            Left            =   1680
            TabIndex        =   20
            Top             =   840
            Width           =   4935
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
            Left            =   -480
            TabIndex        =   31
            Top             =   3840
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
            TabIndex        =   30
            Top             =   3120
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
            TabIndex        =   29
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos"
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
            Index           =   8
            Left            =   0
            TabIndex        =   28
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
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
            Index           =   15
            Left            =   -600
            TabIndex        =   27
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
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
            Index           =   16
            Left            =   0
            TabIndex        =   26
            Top             =   960
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2040
      Index           =   3
      Left            =   5850
      TabIndex        =   3
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
         TabIndex        =   4
         Top             =   120
         Width           =   3615
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   4
            Left            =   1920
            Picture         =   "HistorialVentas.frx":080F
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   240
            Picture         =   "HistorialVentas.frx":0E9A
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "?Desea cancelar el movimiento?"
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
            TabIndex        =   6
            Top             =   120
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
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
            TabIndex        =   10
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   92536833
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   1200
            Width           =   14775
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   4200
            TabIndex        =   12
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   92536833
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
            TabIndex        =   16
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
            TabIndex        =   15
            Top             =   3120
            Width           =   13095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"HistorialVentas.frx":16A9
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
            TabIndex        =   14
            Top             =   3840
            Width           =   14655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"HistorialVentas.frx":1737
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
            TabIndex        =   13
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
            TabIndex        =   11
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
Attribute VB_Name = "frmHistorialVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//OTROS
    Dim i                   As Long
    Dim Prt                 As Printer
    
    '//RECORSET
    Dim Rs                  As New adodb.Recordset  'Cabecera ComprasVentas
    Dim Rs1                 As New adodb.Recordset  'VentasCompras
    Dim Rs2                 As New adodb.Recordset  'Pagos
    Dim Rs3                 As New adodb.Recordset  'inventarios
    Dim Rs4                 As New adodb.Recordset  'puntos
    Dim Rs5                 As New adodb.Recordset  'terjeta
    Dim Rs6                 As New adodb.Recordset  'ticket
    
    '//VENTAS
    Dim c1                  As String
    Dim c2                  As String
    Dim c3                  As String
    Dim c4                  As String
    Dim c5                  As String
    Dim c6                  As String
    Dim c7                  As String
    Dim nc                  As Long
    Dim v2                  As String               'folio
    
    '//CLIENTES
    Dim ClienteMayorista    As String
    Dim v1                  As Long                 'id cliente
    Dim ListaPrecios        As Long
    
    '//TICKET
    Dim vTicketSubtotal     As String
    Dim vTicketIva          As String
    Dim vTicketTotal        As String
        
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        frmHistorialVentas.Pagar.Caption = "Cobrar Venta"
        
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
            
            If StTipoVenta = "Local" Then
                .Open "Select * from PO_HEADERS_ALL_R where LugarVenta = 'Local';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            End If
                
            If StTipoVenta = "Domicilio" Then
                .Open "Select * from PO_HEADERS_ALL_R where LugarVenta = 'Domicilio';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            End If
                
            If StTipoVenta = "Abiertas" Then
                .Open "Select * from PO_HEADERS_ALL_R where totalPagado < total;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
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
            
            With Rs4
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from RA_POINT_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            End With
        End If
        
        With Combo2
            .AddItem "Efectivo"
            .AddItem "Tarjeta"
            .AddItem "Puntos"
            
            .Text = "Efectivo"
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:DTPicker1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub List1_Click()
        On Error GoTo errHandler
        
        Label2.Caption = Get_Comentario(Trim(Mid(List1.Text, 1, 10)))
        
        List2.Clear
        
        
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from PO_LINES_ALL", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
            
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:List1_Click" & vbTab & err.Number & vbTab & err.Description
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
                                
                        With .Sections("Secci?n4")
                            .Controls("Etiqueta2").Caption = "TICKET DE VENTA"
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
                                
                        With .Sections("Secci?n1")
                            .Controls("Texto1").DataField = "cantidad"
                            .Controls("Texto2").DataField = "articulo"
                            .Controls("Texto3").DataField = "subtotal"
                        End With
                                    
                        With .Sections("Secci?n5")
                            .Controls("Etiqueta23").Caption = "$ " & vTicketSubtotal 'subtotal
                            .Controls("Etiqueta26").Caption = "$ " & vTicketIva 'iva
                            .Controls("Etiqueta27").Caption = "$ " & vTicketTotal 'total
                            .Controls("Etiqueta25").Visible = True
                            .Controls("Etiqueta28").Visible = True
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Ticket_Click" & vbTab & err.Number & vbTab & err.Description
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
            
            v2 = Trim(Mid(List1.Text, 1, 10))
            v1 = Get_ClienteFolio(v2)
            
            Text2(0) = Replace(Format(Val(Mid(List1.Text, 68, 12)) - Val(Mid(List1.Text, 81, 12)), "0.00"), ",", ".")
            Text2(1) = ""
            Text2(3) = Get_ClientePuntos(v1) 'Funcion Puntos
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Pagar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Text2_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo2_click()
        On Error GoTo errHandler
        
        If Combo2 = "Puntos" Then
            Text2(1).Enabled = False
        Else
            Text2(1).Enabled = True
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Combo2_click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        
        Frame1.Enabled = True
        
        Frame2(0).Visible = False
            
        'Pago con tarjeta
        If Combo2 = "Tarjeta" Then
            If Text2(4) = "" Then
                MsgBox "La referencia es obligatoria en el pago con tarjeta", vbCritical, "Advertencia"
                
                'frmHistorialVentas.Show 1
                
                Frame1.Enabled = False
                
                Frame2(0).Visible = True
                
                Exit Sub
            End If
                
            If Val(Text2(1)) = 0 Then
                Text2(1) = Text2(0)
            End If
                
            With Rs5
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from RA_BANK_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                .AddNew
                    .Fields(1) = Date
                    .Fields(2) = Text2(1)
                    .Fields(3) = v2
                    .Fields(4) = "No"
                    .Fields(5) = frmMenuInicial.Combo1.Text
                    .Fields(6) = Text2(4)
                    .Fields(7) = v1
                .Update
                .Requery
            End With
        End If
            
        'Pago con puntos
        If Combo2 = "Puntos" Then
            If Val(Text2(3)) > Val(Text2(0)) Then
                Text2(1).Text = Text2(0).Text
            Else
                Text2(1).Text = Text2(3).Text
            End If
                
            If Val(Text2(1)) <> 0 Then
                With Rs4
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from RA_POINT_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                        .Fields(1) = v1
                        .Fields(2) = Replace(Val(Text2(1)) * -1, ",", ".")
                        .Fields(3) = v2
                        .Fields(4) = "No"
                        .Fields(5) = Date
                    .Update
                    .Requery
                End With
            End If
        End If
            
        'pago con efectivo
        If Combo2 = "Efectivo" And Val(Text2(1)) > 0 Then
            With Rs2
                .AddNew
                    .Fields(1) = Date
                    .Fields(2) = "Pago de venta"
                            
                    If Val(Text2(1)) >= Val(Text2(0)) Then
                        .Fields(3) = Replace(Format(Val(Text2(0)), "0.00"), ",", ".")
                    Else
                        .Fields(3) = Replace(Format(Val(Text2(1)), "0.00"), ",", ".")
                    End If
                        
                    .Fields(4) = Trim(Mid(List1.Text, 1, 10))
                    .Fields(5) = "No"
                    .Fields(6) = frmMenuInicial.Combo1.Text
                .Update
                .Requery
            End With
        End If
            
        'si la venta no es credito
        If Val(Text2(1)) > 0 And Combo2 <> "Puntos" Then
            'puntos
            ClienteMayorista = Get_ClienteMayorista(v1)
            ListaPrecios = Get_ClienteListaP(v1)
                If ClienteMayorista = "No" And ListaPrecios = 1 Then
                    With Rs4
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from RA_POINT_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                            .Fields(1) = v1 'id cliente
                            
                            If Val(Text2(1)) > (Val(Text2(0))) Then
                                .Fields(2) = Replace(Round(Val(Text2(0)) * Val(PcValorPuntos), 2), ",", ".")
                            Else
                                .Fields(2) = Replace(Round(Val(Text2(1)) * Val(PcValorPuntos), 2), ",", ".")
                            End If
                            
                            .Fields(3) = v2 'folio
                            .Fields(4) = "No"
                            .Fields(5) = Date
                        .Update
                        .Requery
                    End With
                End If
            End If
            
            If Val(Text2(1)) > 0 Then
                With Rs1
                    Dim DineroRestante As String
                    Dim TMovimiento As String
                    Dim TPagado As String
                    Dim TDebido As String
                    
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
                    
                    '.Update
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
            
            Archivo.Enabled = True
        Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Cancelar_Click()
        On Error GoTo errHandler
        
        If Mid(List1.Text, 1, 10) <> "" Then
            Frame1.Enabled = False
            
            Frame2(3).Visible = True
            
            Archivo.Enabled = False
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Cancelar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command3_Click()
        On Error GoTo errHandler
        
        With Rs1
            .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
            .Requery
            
            If .RecordCount > 0 Then
                .MoveFirst
                
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
            
            If .RecordCount > 0 Then
                .MoveFirst
                
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
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                Do Until .EOF
                    .Fields(9) = "Si"
                    .MoveNext
                Loop
    
                .Filter = ""
                .Requery
            End If
        End With
        
        With Rs4
            .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
            .Requery
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                Do Until .EOF
                    .Fields(4) = "Si"
                    .MoveNext
                Loop
    
                .Filter = ""
                .Requery
            End If
        End With
        
        With Rs5
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from RA_BANK_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
            .Requery
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                Do Until .EOF
                    .Fields(4) = "Si"
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
                        
        Frame1.Enabled = True
        
        Frame2(3).Visible = False
        
        Archivo.Enabled = True
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Command3_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 4
                Frame1.Enabled = True
                
                Frame2(3).Visible = False
                
                Archivo.Enabled = True
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Command1_Click" & vbTab & err.Number & vbTab & err.Description
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
    
            With frmAgregarArticuloVentas
                .Caption = "A?adir art?culos"
                .Text1(0) = Trim(Mid(List1.Text, 1, 10))
                .Text1(7) = Rs1.Fields(2).Value
                
                If IsNull(Rs1.Fields(4).Value) = False Then
                    .Text1(8) = Rs1.Fields(4).Value
                End If
                
                .Show 1
            End With
            
            Unload frmHistorialVentas
            
            Set frmHistorialVentas = Nothing
            
            frmHistorialVentas.Show
        Else
            MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Agregar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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
        If Rs4.State = 1 Then Rs4.Close
        If Rs5.State = 1 Then Rs5.Close
        If Rs6.State = 1 Then Rs6.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Rs2 = Nothing
        Set Rs3 = Nothing
        Set Rs4 = Nothing
        Set Rs5 = Nothing
        Set Rs6 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmHistorialVentas:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
