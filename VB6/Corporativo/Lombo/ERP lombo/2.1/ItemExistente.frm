VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmItemExistente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Editar Artículos"
   ClientHeight    =   7260
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   17325
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
   ScaleHeight     =   7260
   ScaleWidth      =   17325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17055
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16815
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4695
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   16335
            _ExtentX        =   28813
            _ExtentY        =   8281
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            HeadLines       =   3
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
            Caption         =   "Listado de artículos"
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
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "ItemExistente.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   5200
            Picture         =   "ItemExistente.frx":07EE
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   2
            Left            =   10160
            Picture         =   "ItemExistente.frx":101C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   3
            Left            =   15120
            Picture         =   "ItemExistente.frx":18B8
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   5880
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   15135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar"
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
            Left            =   240
            TabIndex        =   2
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
Attribute VB_Name = "frmItemExistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs  As New adodb.Recordset
    
    '//OTROS
    Dim i   As Long
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_SYSTEM_ITEMS order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        If Rs.RecordCount > 0 Then
            Set DataGrid1.DataSource = Rs
            
            With DataGrid1
                For i = 0 To 12
                    .Columns(i).Width = 1160
                Next i
                
                .Columns(1).Width = 2500
                .Columns(2).Width = 6000
                .Columns(10).Width = 2500
                
                .Columns(3).Alignment = dbgRight
                .Columns(4).Alignment = dbgRight
                .Columns(5).Alignment = dbgRight
                .Columns(6).Alignment = dbgRight
                .Columns(7).Alignment = dbgRight
                .Columns(8).Alignment = dbgRight
                .Columns(11).Alignment = dbgRight
                
                .Columns(0).Visible = False
                .Columns(1).Locked = True
                .Columns(2).Locked = True
                .Columns(9).Locked = True
                .Columns(10).Locked = True
                .Columns(7).Caption = "Precio5 o precio de compra"
                .Columns(10).Caption = "Categoria"
                .Columns(12).Caption = "Tipo"
            End With
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
            
            Unload Me
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemExistente:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
                    For i = 0 To 12
                        .Columns(i).Width = 1160
                    Next i
                    
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 6000
                    .Columns(10).Width = 2500
                    
                    .Columns(3).Alignment = dbgRight
                    .Columns(4).Alignment = dbgRight
                    .Columns(5).Alignment = dbgRight
                    .Columns(6).Alignment = dbgRight
                    .Columns(7).Alignment = dbgRight
                    .Columns(8).Alignment = dbgRight
                    .Columns(11).Alignment = dbgRight
                    
                    .Columns(0).Visible = False
                    .Columns(1).Locked = True
                    .Columns(2).Locked = True
                    .Columns(5).Locked = True
                    .Columns(6).Locked = True
                    
                    .Columns(10).Caption = "Categoria"
                    .Columns(12).Caption = "Tipo"
                    .Columns(7).Caption = "Precio5 o precio de compra"
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemExistente:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                Rs.MoveFirst
            
            Case 1
                If Rs.BOF = False Then Rs.MovePrevious
            
            Case 2
                If Rs.EOF = False Then Rs.MoveNext
            
            Case 3
                Rs.MoveLast
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemExistente:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        With Rs
            .Update
            .Requery
        End With
        
        With DataGrid1
            For i = 0 To 12
                .Columns(i).Width = 1160
            Next i
                
            .Columns(1).Width = 2500
            .Columns(2).Width = 6000
            .Columns(10).Width = 2500
            
            .Columns(3).Alignment = dbgRight
            .Columns(4).Alignment = dbgRight
            .Columns(5).Alignment = dbgRight
            .Columns(6).Alignment = dbgRight
            .Columns(7).Alignment = dbgRight
            .Columns(8).Alignment = dbgRight
            .Columns(11).Alignment = dbgRight
                
            .Columns(0).Visible = False
            .Columns(1).Locked = True
            .Columns(2).Locked = True
            .Columns(5).Locked = True
            .Columns(6).Locked = True
            
            .Columns(10).Caption = "Categoria"
            .Columns(12).Caption = "Tipo"
            .Columns(7).Caption = "Precio5 o precio de compra"
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemExistente:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        
        Rs.Requery
        
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemExistente:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemExistente:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
