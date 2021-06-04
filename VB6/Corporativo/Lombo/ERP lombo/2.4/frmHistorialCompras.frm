VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHistorialCompras 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Historial de compras"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   3240
      Index           =   0
      Left            =   4170
      TabIndex        =   3
      Top             =   2917
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3015
         Index           =   2
         Left            =   120
         TabIndex        =   4
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
            TabIndex        =   20
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            ForeColor       =   &H00E0E0E0&
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
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
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
            TabIndex        =   6
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
               Name            =   "MS Sans Serif"
               Size            =   13.5
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
            TabIndex        =   5
            Top             =   120
            Width           =   5055
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
            TabIndex        =   10
            Top             =   1680
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
            TabIndex        =   9
            Top             =   960
            Width           =   2055
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
            TabIndex        =   8
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17175
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   8655
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16935
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   0
            Left            =   1320
            TabIndex        =   13
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   4210752
            CalendarForeColor=   14737632
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   14737632
            CalendarTrailingForeColor=   8421504
            Format          =   130088961
            CurrentDate     =   43915
            MaxDate         =   2958101
         End
         Begin VB.ListBox List2 
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
            Height          =   3480
            Left            =   120
            TabIndex        =   12
            Top             =   5040
            Width           =   16695
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
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   16695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   5160
            TabIndex        =   15
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   4210752
            CalendarForeColor=   14737632
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   14737632
            CalendarTrailingForeColor=   8421504
            Format          =   130088961
            CurrentDate     =   43915
            MaxDate         =   2958101
         End
         Begin VB.Label Label1 
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
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   3720
            Width           =   1935
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
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
            Height          =   615
            Left            =   2280
            TabIndex        =   18
            Top             =   3720
            Width           =   14535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "ARTICULO                                        CANTIDAD         PRECIO        SUBTOTAL            IVA               TOTAL"
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
            Left            =   240
            TabIndex        =   17
            Top             =   4680
            Width           =   12855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmHistorialCompras.frx":0000
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
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   14655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "HASTA"
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
            Left            =   4080
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DESDE"
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
'***********************************************************************************
'Nombre:        frmHistorialCompras
'Proposito:     Consultar, Cancelar, Modificar y Pagar compras previamente
'               registradas
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        25/05/2021     Alfredo Hernandez    Se agrego usuario, fechas de
'                                               creacion, modificacion a los insert
'                                               y updates
'
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    '//OTROS
    Dim i                   As Long
    Dim Prt                 As Printer
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset  'Cabecera ComprasVentas
    Dim RS1                 As New adodb.Recordset  'VentasCompras
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
    Dim sql                 As String
    '//PAGOS
    Dim DineroRestante      As String
    Dim TMovimiento         As String
    Dim TPagado             As String
    Dim TDebido             As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        For i = 0 To 1
            With DTPicker1(i)
                .Value = Date
            End With
        Next i
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            With frmHistorialCompras
                With .Pagar
                    .Caption = "Pagar Compra"
                End With
            End With
            
            If StTipoCompra = "Pagadas" Then
                .Open "Select * from PO_HEADERS_ALL_P where totalPagado = total;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            End If
            
            If StTipoCompra = "No Pagadas" Then
                .Open "Select * from PO_HEADERS_ALL_P where totalPagado < total;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            End If
            .Requery
            If .RecordCount > 0 Then
                .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                .Requery
                With List1
                    .Clear
                End With
                
                With List2
                    .Clear
                End With
                
                Do Until .EOF
                    c1 = Mid(Rs!folio, 1, 10)
                    c2 = Mid(Rs!fecha, 1, 10)
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
                    With List1
                        .AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                    End With
                    .MoveNext
                Loop
                
                With RS1
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
        End With
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
                    With List1
                        .Clear
                    End With
                    
                    With List2
                        .Clear
                    End With
                    
                    Do Until .EOF
                        c1 = Mid(Rs!folio, 1, 10)
                        c2 = Mid(Rs!fecha, 1, 10)
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
                        
                        With List1
                            .AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                        End With
                        .MoveNext
                    Loop
                End With
            Case 1
                With Rs
                    .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                    .Requery
                    With List1
                        .Clear
                    End With
                    
                    With List2
                        .Clear
                    End With
                    
                    Do Until .EOF
                        c1 = Mid(Rs!folio, 1, 10)
                        c2 = Mid(Rs!fecha, 1, 10)
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
                        
                        With List1
                            .AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                        End With
                        .MoveNext
                    Loop
                End With
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
        With Label2
            .Caption = Get_Comentario(Trim(Mid(List1.Text, 1, 10)))
        End With
        
        With List2
            .Clear
        End With
        
        With RS1
            .Requery
            .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
            Do Until .EOF
                c1 = Mid(RS1!DescripcionArticulo & " (" & RS1!UDM & ")" & " (" & RS1!CodigoArticulo & ")", 1, 27)
                c2 = Replace(Format(Mid(RS1!cantidad, 1, 12), "0.00"), ",", ".")
                c3 = Replace(Format(Mid(RS1!Precio, 1, 12), "0.00"), ",", ".")
                c4 = Replace(Format(Mid(RS1!Subtotal, 1, 12), "0.00"), ",", ".")
                c5 = Replace(Format(Mid(RS1!IVA, 1, 12), "0.00"), ",", ".")
                c6 = Replace(Format(Mid(RS1!Total, 1, 12), "0.00"), ",", ".")
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
                
                With List2
                    .AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5 & " " & c6
                End With
                .MoveNext
            Loop
        End With
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
        With List1
            If Mid(.Text, 1, 10) <> "" Then
                With Archivo
                    .Enabled = False
                End With
                
                With Frame1
                    .Enabled = False
                End With
                
                With Frame2(0)
                    .Visible = True
                End With
                
                With RS1
                    .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                    .Requery
                End With
                Text2(0) = Replace(Format(Val(Mid(.Text, 68, 12)) - Val(Mid(.Text, 81, 12)), "0.00"), ",", ".")
            Else
                MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
            End If
        End With
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
            With Frame1
                .Enabled = True
            End With
            
            With Frame2(0)
                .Visible = False
            End With
            
            With Text2(1)
                If Val(.Text) > 0 Then
                    With Rs2
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
                                .Value = Trim(Mid(List1.Text, 1, 10))                                   'folio
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
                    
                    With Text2(1)
                        DineroRestante = Val(.Text)
                    End With
                    
                    With RS1
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
                    
                    With List1
                        .Clear
                    End With
                    
                    With List2
                        .Clear
                    End With
                    
                    With Rs
                        .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                        .Requery
                        Do Until .EOF
                            c1 = Mid(Rs!folio, 1, 10)
                            c2 = Mid(Rs!fecha, 1, 10)
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
                            
                            With List1
                                .AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                            End With
                            .MoveNext
                        Loop
                    End With
                    
                    With Frame1
                        .Enabled = True
                    End With
                    
                    With Frame2(0)
                        .Visible = False
                    End With
                End If
            End With
        End If
        
        With Archivo
            .Enabled = True
        End With
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
                sql = "UPDATE PO_LINES_ALL SET CANCELADO = 'Si', LAST_UPDATE_DATE = '" & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS") & "', LAST_UPDATED_BY = '" & StUsuario & "' WHERE Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                With Cn
                    .Execute sql
                End With
                sql = "UPDATE RA_CASH_TRANSACTIONS SET CANCELADO = 'Si', LAST_UPDATE_DATE = '" & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS") & "', LAST_UPDATED_BY = '" & StUsuario & "' WHERE Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                With Cn
                    .Execute sql
                End With
                sql = "UPDATE MTL_MATERIAL_TRANSACTIONS SET CANCELADO = 'Si', LAST_UPDATE_DATE = '" & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS") & "', LAST_UPDATED_BY = '" & StUsuario & "' WHERE Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                With Cn
                    .Execute sql
                End With
                MsgBox "Movimiento Cancelado", vbOKOnly, "Terminado"
                With List1
                    .Clear
                End With
                
                With List2
                    .Clear
                End With
                
                With Rs
                    .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                    .Requery
                    Do Until Rs.EOF
                        c1 = Mid(Rs!folio, 1, 10)
                        c2 = Mid(Rs!fecha, 1, 10)
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
                        
                        With List1
                            .AddItem c1 & " " & c2 & " " & c3 & " " & c5 & " " & c6 & " " & c7
                        End With
                        .MoveNext
                    Loop
                End With
                
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
        With List1
            If Mid(.Text, 1, 10) <> "" Then
                With RS1
                    .Filter = "Folio = '" & Trim(Mid(List1.Text, 1, 10)) & "'"
                    .Requery
                End With
                
                With frmAgregarArticuloCompras
                    .Caption = "Añadir artículos"
                    With .Text1(0)
                        .Text = Trim(Mid(List1.Text, 1, 10))
                    End With
                    
                    With .Text1(7)
                        .Text = RS1.Fields(2).Value
                    End With
                    
                    If IsNull(RS1.Fields(4).Value) = False Then
                        With .Text1(8)
                            .Text = RS1.Fields(4).Value
                        End With
                    End If
                    .Show 1
                End With
                Unload frmHistorialCompras
                Set frmHistorialCompras = Nothing
                
                With frmHistorialCompras
                    .Show
                End With
            Else
                MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
            End If
        End With
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
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set RS1 = Nothing
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
