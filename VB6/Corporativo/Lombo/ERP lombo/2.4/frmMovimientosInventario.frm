VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMovimientosInventario 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Movimientos de inventario"
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
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17175
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   8535
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   16935
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
            Index           =   1
            Left            =   2520
            TabIndex        =   5
            Top             =   1720
            Width           =   14295
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "EXCEL"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   7920
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
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
            Height          =   465
            Left            =   2520
            TabIndex        =   6
            Top             =   1200
            Width           =   14295
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
            Index           =   0
            Left            =   2520
            TabIndex        =   4
            Top             =   720
            Width           =   14295
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   0
            Left            =   2520
            TabIndex        =   1
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
            Format          =   140771329
            CurrentDate     =   43915
            MaxDate         =   2958101
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   6360
            TabIndex        =   2
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
            Format          =   140771329
            CurrentDate     =   43915
            MaxDate         =   2958101
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4815
            Left            =   120
            TabIndex        =   11
            Top             =   2880
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   8493
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421504
            BorderStyle     =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   14737632
            HeadLines       =   2
            RowHeight       =   28
            RowDividerStyle =   5
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
                  LCID            =   3082
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
                  LCID            =   3082
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
            Index           =   5
            Left            =   360
            TabIndex        =   14
            Top             =   1725
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMovimientosInventario.frx":0000
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
            Left            =   480
            TabIndex        =   12
            Top             =   2520
            Width           =   15975
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
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   975
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
            Left            =   5280
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACCION"
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
            Left            =   360
            TabIndex        =   8
            Top             =   1200
            Width           =   1935
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
            Index           =   1
            Left            =   480
            TabIndex        =   7
            Top             =   720
            Width           =   1815
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Exportar 
         Caption         =   "Exportar a Excel"
         Shortcut        =   ^E
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmMovimientosInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmMovimientosInventario
'Proposito:     Consulta de entradas/salidas de artículos
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset
Dim RS1 As New adodb.Recordset
'//OTROS
Dim i As Long

Private Sub Form_Load()
    On Error GoTo errHandler
    With DTPicker1(0)
        .Value = Date
    End With

    With DTPicker1(1)
        .Value = Date
    End With

    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With Rs
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from MTL_MATERIAL_TRANSACTIONS_V order by 1,2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        If .RecordCount > 0 Then
            .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
            .Requery
            With DataGrid1
                Set .DataSource = Rs

                With .Columns(0)
                    .Width = 1700
                    .Locked = True
                End With

                With .Columns(1)
                    .Width = 2000
                    .Locked = True
                End With

                With .Columns(2)
                    .Width = 4400
                    .Locked = True
                End With

                With .Columns(3)
                    .Width = 3000
                    .Locked = True
                End With

                With .Columns(4)
                    .Width = 1300
                    .Locked = True
                    .Alignment = dbgRight
                End With

                With .Columns(5)
                    .Width = 1200
                    .Locked = True
                End With

                With .Columns(6)
                    .Visible = False
                End With
            End With
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
        End If
    End With
    With RS1
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select distinct transaccion as transaccion from MTL_MATERIAL_TRANSACTIONS_V order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        If .RecordCount <> 0 Then
            .MoveFirst
            With Combo1
                .AddItem ""
            End With

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

            With .Columns(0)
                .Width = 1700
                .Locked = True
            End With

            With .Columns(1)
                .Width = 2000
                .Locked = True
            End With

            With .Columns(2)
                .Width = 4400
                .Locked = True
            End With

            With .Columns(3)
                .Width = 3000
                .Locked = True
            End With

            With .Columns(4)
                .Width = 1300
                .Locked = True
                .Alignment = dbgRight
            End With

            With .Columns(5)
                .Width = 1200
                .Locked = True
            End With

            With .Columns(6)
                .Visible = False
            End With
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

            With .Columns(0)
                .Width = 1700
                .Locked = True
            End With

            With .Columns(1)
                .Width = 2000
                .Locked = True
            End With

            With .Columns(2)
                .Width = 4400
                .Locked = True
            End With

            With .Columns(3)
                .Width = 3000
                .Locked = True
            End With

            With .Columns(4)
                .Width = 1300
                .Locked = True
                .Alignment = dbgRight
            End With

            With .Columns(5)
                .Width = 1200
                .Locked = True
            End With

            With .Columns(6)
                .Visible = False
            End With
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

            With .Columns(0)
                .Width = 1700
                .Locked = True
            End With

            With .Columns(1)
                .Width = 2000
                .Locked = True
            End With

            With .Columns(2)
                .Width = 4400
                .Locked = True
            End With

            With .Columns(3)
                .Width = 3000
                .Locked = True
            End With

            With .Columns(4)
                .Width = 1300
                .Locked = True
                .Alignment = dbgRight
            End With

            With .Columns(5)
                .Width = 1200
                .Locked = True
            End With

            With .Columns(6)
                .Visible = False
            End With
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

        With .Columns(0)
            .Width = 1700
            .Locked = True
        End With

        With .Columns(1)
            .Width = 2000
            .Locked = True
        End With

        With .Columns(2)
            .Width = 4400
            .Locked = True
        End With

        With .Columns(3)
            .Width = 3000
            .Locked = True
        End With

        With .Columns(4)
            .Width = 1300
            .Locked = True
            .Alignment = dbgRight
        End With

        With .Columns(5)
            .Width = 1200
            .Locked = True
        End With

        With .Columns(6)
            .Visible = False
        End With
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

        With .Columns(0)
            .Width = 1700
            .Locked = True
        End With

        With .Columns(1)
            .Width = 2000
            .Locked = True
        End With

        With .Columns(2)
            .Width = 4400
            .Locked = True
        End With

        With .Columns(3)
            .Width = 3000
            .Locked = True
        End With

        With .Columns(4)
            .Width = 1300
            .Locked = True
            .Alignment = dbgRight
        End With

        With .Columns(5)
            .Width = 1200
            .Locked = True
        End With

        With .Columns(6)
            .Visible = False
        End With
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

Private Sub Command2_Click()
    On Error GoTo errHandler
    'PARA EXPORTAR A EXCEL
    Dim N As Long, sTemp As String
    Dim FileName As String
    FileName = App.Path & "\Temp\TEMP_MI_" & CStr(Format(Date, "YYYYMMDD")) & "_" & CStr(Format(Time, "HHMMSS")) & ".xls"
    Open FileName For Output As #1
    'ENCABEZADO
    sTemp = "INFORME DE MOVIMIENTOS DE INVENTARIO"
    Print #1, sTemp
    sTemp = vbNullString
    sTemp = "Desde: " & DTPicker1(0).Value & " Hasta: " & DTPicker1(1).Value
    Print #1, sTemp
    sTemp = vbNullString
    sTemp = "Articulo: " & Text1(0).Text
    Print #1, sTemp
    sTemp = vbNullString
    sTemp = "Tipo de transacciòn: " & Combo1.Text
    Print #1, sTemp
    sTemp = vbNullString
    sTemp = "Fecha de Ejecucion del informe: " & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
    Print #1, sTemp
    sTemp = vbNullString
    Print #1, sTemp
    sTemp = vbNullString
    'CABECERA
    For N = 0 To Rs.Fields.Count - 1
        sTemp = sTemp & UCase(Rs.Fields(N).Name) & IIf(N = Rs.Fields.Count - 1, vbNullString, vbTab)
    Next N

    Print #1, sTemp
    sTemp = vbNullString
    'DETALLE
    With Rs
        .MoveFirst
        Do Until .EOF
            For N = 0 To .Fields.Count - 1
                If N = 4 Then    'CONVERTIR A NUMERO
                    sTemp = sTemp & Replace(CStr(.Fields(N).Value), ",", ".") & IIf(N = .Fields.Count - 1, vbNullString, vbTab)
                Else
                    sTemp = sTemp & .Fields(N).Value & IIf(N = .Fields.Count - 1, vbNullString, vbTab)
                End If
            Next N

            Print #1, sTemp
            sTemp = vbNullString
            .MoveNext
        Loop
    End With

    Close #1

    'PARA ABRIR EL ARCHIVO DE EXCEL AL TERMINAR DE EXPORTAR
    Dim xltmp As Excel.Application

    Set xltmp = New Excel.Application

    With xltmp
        With .Workbooks
            .Open FileName
        End With

        With .Range("A7", "H7")
            With .Interior
                .Color = RGB(80, 80, 80)
            End With

            With .Font
                .Color = RGB(255, 255, 255)
            End With
        End With

        With .ActiveWorkbook
            .Save
        End With
        .Visible = True
    End With
    Unload Me
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMovimientosInventario:Command2_Click" & vbTab & err.Number & vbTab & err.Description
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
    With DataGrid1
        Set .DataSource = Nothing
    End With

    With Rs
        If .State = 1 Then .Close
    End With

    With RS1
        If .State = 1 Then .Close
    End With

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set RS1 = Nothing
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
