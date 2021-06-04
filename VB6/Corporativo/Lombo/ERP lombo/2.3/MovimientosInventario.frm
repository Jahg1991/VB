VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMovimientosInventario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Movimientos de inventario"
   ClientHeight    =   7665
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
   Moveable        =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13455
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1800
            TabIndex        =   5
            Top             =   1320
            Width           =   11415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   1800
            TabIndex        =   4
            Top             =   720
            Width           =   11415
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4935
            Left            =   240
            TabIndex        =   7
            Top             =   1800
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   8705
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            HeadLines       =   2
            RowHeight       =   28
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Movimientos de Inventario"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   0
            Left            =   1800
            TabIndex        =   2
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   131268609
            CurrentDate     =   43915
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   4680
            TabIndex        =   3
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   131268609
            CurrentDate     =   43915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transacción"
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
            TabIndex        =   10
            Top             =   1440
            Width           =   1455
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
            Index           =   1
            Left            =   720
            TabIndex        =   9
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Width           =   975
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
            Left            =   3840
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmMovimientosInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs As New adodb.Recordset
    Dim Rs1 As New adodb.Recordset
    
    '//OTROS
    Dim i As Long

    Private Sub Form_Load()
        On Error GoTo errHandler
        
        DTPicker1(0).Value = Date
        DTPicker1(1).Value = Date
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_MATERIAL_TRANSACTIONS_V order by 1,2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        If Rs.RecordCount > 0 Then
            Rs.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
            Rs.Requery
        
            Set DataGrid1.DataSource = Rs
            
            With DataGrid1
                .Columns(0).Width = 1500
                .Columns(1).Width = 2000
                .Columns(2).Width = 4000
                .Columns(3).Width = 3000
                .Columns(4).Width = 1300
                .Columns(5).Width = 1200
                
                .Columns(0).Locked = True
                .Columns(1).Locked = True
                .Columns(2).Locked = True
                .Columns(3).Locked = True
                .Columns(4).Locked = True
                .Columns(5).Locked = True
                
                .Columns(6).Visible = False
            End With
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"

            'Unload Me
        End If
        
         With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select distinct transaccion as transaccion from MTL_MATERIAL_TRANSACTIONS_V order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            
            If .RecordCount <> 0 Then
                .MoveFirst
                
                Combo1.AddItem ""
                
                While Not .EOF
                    Combo1.AddItem .Fields(0).Value
                    .MoveNext
                Wend
                
                .Close
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub DTPicker1_Change(Index As Integer)
        On Error Resume Next
        
        Select Case Index
            Case 0
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "  SELECT *                                                                                          " & _
                          "    FROM MTL_MATERIAL_TRANSACTIONS_V                                                                " & _
                          "   WHERE (Codigo LIKE ISNULL('%" & Text1(0) & "%',Codigo)                                           " & _
                          "      OR Descripcion LIKE ISNULL('%" & Text1(0) & "%',Descripcion))                                 " & _
                          "     AND Transaccion = CASE WHEN '" & Combo1 & "' = '' THEN Transaccion ELSE '" & Combo1 & "' END   " & _
                          "ORDER BY 1,2;                                                                                       ", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                    .Requery
                End With
                
                With DataGrid1
                    Set .DataSource = Rs
                    
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 4000
                    .Columns(3).Width = 3000
                    .Columns(4).Width = 1300
                    .Columns(5).Width = 1200
                    
                    .Columns(0).Locked = True
                    .Columns(1).Locked = True
                    .Columns(2).Locked = True
                    .Columns(3).Locked = True
                    .Columns(4).Locked = True
                    .Columns(5).Locked = True
                    
                    .Columns(6).Visible = False
                End With
                
            Case 1
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "  SELECT *                                                                                          " & _
                          "    FROM MTL_MATERIAL_TRANSACTIONS_V                                                                " & _
                          "   WHERE (Codigo LIKE ISNULL('%" & Text1(0) & "%',Codigo)                                           " & _
                          "      OR Descripcion LIKE ISNULL('%" & Text1(0) & "%',Descripcion))                                 " & _
                          "     AND Transaccion = CASE WHEN '" & Combo1 & "' = '' THEN Transaccion ELSE '" & Combo1 & "' END   " & _
                          "ORDER BY 1,2;                                                                                       ", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                    .Requery
                End With
                
                With DataGrid1
                    Set .DataSource = Rs
                    
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 4000
                    .Columns(3).Width = 3000
                    .Columns(4).Width = 1300
                    .Columns(5).Width = 1200
                    
                    .Columns(0).Locked = True
                    .Columns(1).Locked = True
                    .Columns(2).Locked = True
                    .Columns(3).Locked = True
                    .Columns(4).Locked = True
                    .Columns(5).Locked = True
                    
                    .Columns(6).Visible = False
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:DTPicker1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Text1_Change(Index As Integer)
        On Error Resume Next
        
        Select Case Index
            Case 0
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "  SELECT *                                                                                          " & _
                          "    FROM MTL_MATERIAL_TRANSACTIONS_V                                                                " & _
                          "   WHERE (Codigo LIKE ISNULL('%" & Text1(0) & "%',Codigo)                                           " & _
                          "      OR Descripcion LIKE ISNULL('%" & Text1(0) & "%',Descripcion))                                 " & _
                          "     AND Transaccion = CASE WHEN '" & Combo1 & "' = '' THEN Transaccion ELSE '" & Combo1 & "' END   " & _
                          "ORDER BY 1,2;                                                                                       ", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                    .Requery
                End With
                
                With DataGrid1
                    Set .DataSource = Rs
                    
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 4000
                    .Columns(3).Width = 3000
                    .Columns(4).Width = 1300
                    .Columns(5).Width = 1200
                            
                    .Columns(0).Locked = True
                    .Columns(1).Locked = True
                    .Columns(2).Locked = True
                    .Columns(3).Locked = True
                    .Columns(4).Locked = True
                    .Columns(5).Locked = True
                            
                    .Columns(6).Visible = False
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
        Private Sub Combo1_Click()
        On Error GoTo errHandler
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "  SELECT *                                                                                          " & _
                  "    FROM MTL_MATERIAL_TRANSACTIONS_V                                                                " & _
                  "   WHERE (Codigo LIKE ISNULL('%" & Text1(0) & "%',Codigo)                                           " & _
                  "      OR Descripcion LIKE ISNULL('%" & Text1(0) & "%',Descripcion))                                 " & _
                  "     AND Transaccion = CASE WHEN '" & Combo1 & "' = '' THEN Transaccion ELSE '" & Combo1 & "' END   " & _
                  "ORDER BY 1,2;                                                                                       ", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
            .Requery
        End With
        
        With DataGrid1
            Set .DataSource = Rs
            
            .Columns(0).Width = 1500
            .Columns(1).Width = 2000
            .Columns(2).Width = 4000
            .Columns(3).Width = 3000
            .Columns(4).Width = 1300
            .Columns(5).Width = 1200
                    
            .Columns(0).Locked = True
            .Columns(1).Locked = True
            .Columns(2).Locked = True
            .Columns(3).Locked = True
            .Columns(4).Locked = True
            .Columns(5).Locked = True
                    
            .Columns(6).Visible = False
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
        On Error GoTo errHandler
        
        Static cadena As String
        
        With Combo1
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
            Else
                 .BackColor = COLOR_NORMAL
            End If
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "  SELECT *                                                                                          " & _
                  "    FROM MTL_MATERIAL_TRANSACTIONS_V                                                                " & _
                  "   WHERE (Codigo LIKE ISNULL('%" & Text1(0) & "%',Codigo)                                           " & _
                  "      OR Descripcion LIKE ISNULL('%" & Text1(0) & "%',Descripcion))                                 " & _
                  "     AND Transaccion = CASE WHEN '" & Combo1 & "' = '' THEN Transaccion ELSE '" & Combo1 & "' END   " & _
                  "ORDER BY 1,2;                                                                                       ", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
            .Requery
        End With
        
        With DataGrid1
            Set .DataSource = Rs
            
            .Columns(0).Width = 1500
            .Columns(1).Width = 2000
            .Columns(2).Width = 4000
            .Columns(3).Width = 3000
            .Columns(4).Width = 1300
            .Columns(5).Width = 1200
                    
            .Columns(0).Locked = True
            .Columns(1).Locked = True
            .Columns(2).Locked = True
            .Columns(3).Locked = True
            .Columns(4).Locked = True
            .Columns(5).Locked = True
                    
            .Columns(6).Visible = False
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        Set DataGrid1.DataSource = Nothing
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
