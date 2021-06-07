VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStock 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Cantidad en mano"
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
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   17175
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   8535
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16935
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
            TabIndex        =   4
            Top             =   7920
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   1440
            TabIndex        =   2
            Top             =   240
            Width           =   15375
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6495
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   11456
            _Version        =   393216
            BackColor       =   8421504
            BorderStyle     =   0
            ColumnHeaders   =   0   'False
            ForeColor       =   14737632
            HeadLines       =   1
            RowHeight       =   28
            RowDividerStyle =   5
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
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   6
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
            BeginProperty Column02 
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
            BeginProperty Column03 
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
            BeginProperty Column04 
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
            BeginProperty Column05 
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
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "UDM"
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
            Left            =   12840
            TabIndex        =   11
            Top             =   840
            Width           =   1995
         End
         Begin VB.Label Label1 
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
            Index           =   5
            Left            =   9840
            TabIndex        =   10
            Top             =   840
            Width           =   3000
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION"
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
            Left            =   4800
            TabIndex        =   9
            Top             =   840
            Width           =   4995
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CODIGO"
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
            Left            =   2880
            TabIndex        =   8
            Top             =   840
            Width           =   1995
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "LOTE"
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
            Left            =   14760
            TabIndex        =   7
            Top             =   840
            Width           =   1500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CATEGORIA"
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
            TabIndex        =   6
            Top             =   840
            Width           =   2500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmStock
'Proposito:     Consulta de existencias de inventario
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
'//OTROS
Dim i As Long

Private Sub Form_Load()
    On Error GoTo errHandler
    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With Rs
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from MTL_ON_HAND_QUANTITIES order by 1,3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        If .RecordCount > 0 Then
            With DataGrid1
                Set .DataSource = Rs

                With .Columns(0)
                    .Locked = True
                End With

                With .Columns(1)
                    .Width = 2500
                    .Locked = True
                    .Visible = False
                End With

                With .Columns(2)
                    .Width = 2000
                    .Locked = True
                End With

                With .Columns(3)
                    .Width = 5000
                    .Locked = True
                End With

                With .Columns(4)
                    .Width = 3000
                    .Locked = True
                    .Alignment = dbgRight
                End With

                With .Columns(5)
                    .Width = 2000
                    .Locked = True
                End With

                With .Columns(6)
                    .Width = 1500
                    .Locked = True
                End With
            End With
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmStock:Form_Load" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Rs
            If Text1(0) = "" Then
                .Filter = ""
                .Requery
            Else
                .Filter = "Codigo like '*" & Text1(0) & "*' or Descripcion like '*" & Text1(0) & "*' or Tipo like '*" & Text1(0) & "*'"
                .Requery
            End If
        End With

        With DataGrid1
            Set .DataSource = Rs
            With .Columns(0)
                .Locked = True
            End With

            With .Columns(1)
                .Width = 2500
                .Locked = True
                .Visible = False
            End With

            With .Columns(2)
                .Width = 2000
                .Locked = True
            End With

            With .Columns(3)
                .Width = 5000
                .Locked = True
            End With

            With .Columns(4)
                .Width = 3000
                .Locked = True
                .Alignment = dbgRight
            End With

            With .Columns(5)
                .Width = 2000
                .Locked = True
            End With

            With .Columns(6)
                .Width = 1500
                .Locked = True
            End With
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmStock:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
    'PARA EXPORTAR A EXCEL
    Dim N As Long, sTemp As String
    Dim FileName As String
    FileName = App.Path & "\Temp\TEMP_STOCK_" & CStr(Format(Date, "YYYYMMDD")) & "_" & CStr(Format(Time, "HHMMSS")) & ".xls"
    Open FileName For Output As #1
    'ENCABEZADO
    sTemp = "INFORME DE CANTIDAD EN MANO"
    Print #1, sTemp
    sTemp = vbNullString
    With Text1(0)
        sTemp = "Filtro: " & .Text
    End With

    Print #1, sTemp
    sTemp = vbNullString
    sTemp = "Fecha de Ejecucion del informe: " & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
    Print #1, sTemp
    sTemp = vbNullString
    Print #1, sTemp
    sTemp = vbNullString
    With Rs
        'CABECERA
        For N = 0 To .Fields.Count - 1
            sTemp = sTemp & UCase(.Fields(N).Name) & IIf(N = .Fields.Count - 1, vbNullString, vbTab)
        Next N

        Print #1, sTemp
        sTemp = vbNullString
        'DETALLE
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

        With .Range("A5", "G5")
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmStock:Command2_Click" & vbTab & err.Number & vbTab & err.Description
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmStock:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set Cn = Nothing
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmStock:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
