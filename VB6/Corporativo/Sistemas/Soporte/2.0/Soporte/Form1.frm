VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitudes"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18135
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   18135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Agenda"
      Height          =   375
      Index           =   6
      Left            =   15240
      TabIndex        =   11
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar solicitudes"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar estatus"
      Height          =   375
      Index           =   5
      Left            =   12240
      TabIndex        =   9
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detalles de la solicitud"
      Height          =   375
      Index           =   4
      Left            =   15240
      TabIndex        =   8
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ùltima solicitud"
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   7
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente solicitud"
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solicitud anterior"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Primera solicitud"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9600
      Top             =   8160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":72FA
      OLEDBString     =   $"Form1.frx":7387
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"Form1.frx":7414
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":749E
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "TIPO"
         Caption         =   "TIPO"
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
      BeginProperty Column02 
         DataField       =   "DETALLES"
         Caption         =   "DETALLES"
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
      BeginProperty Column03 
         DataField       =   "FECHA"
         Caption         =   "FECHA"
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
      BeginProperty Column04 
         DataField       =   "HORA"
         Caption         =   "HORA"
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
      BeginProperty Column05 
         DataField       =   "ESTATUS"
         Caption         =   "ESTATUS"
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
      BeginProperty Column06 
         DataField       =   "SOLICITANTE"
         Caption         =   "SOLICITANTE"
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
      BeginProperty Column07 
         DataField       =   "ATENDIDO"
         Caption         =   "ATENDIDO"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4965.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2684.977
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2805.166
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Resueltas"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "En proceso"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Pendientes"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
' Declaración del Api GetUserName
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long) As Long
    
    ' Retorna un String con el nombre de usuario actual de windows
' ************************************************************
Public Function get_Usuario() As String
      
    Dim Nombre As String, ret As Long
      
    ' Buffer
    Nombre = Space$(250)
      
    ' Tamaño
    ret = Len(Nombre)
      
    If GetUserName(Nombre, ret) = 0 Then
        get_Usuario = vbNullString
    Else
        ' Extrae solo los caracteres
        get_Usuario = Left$(Nombre, ret - 1)
    End If
      
End Function


Private Sub Command1_Click(Index As Integer)
    
    On Error Resume Next
    Select Case Index
        Case 0
            Adodc1.Recordset.MoveFirst
        Case 1
            Adodc1.Recordset.MovePrevious
        Case 2
            Adodc1.Recordset.MoveNext
        Case 3
            Adodc1.Recordset.MoveLast
        Case 4
            Form2.Show
            Form2.Label3.Caption = DataGrid1.Columns(0).Value
            Form2.Adodc1.Recordset.Filter = "ID like " + Form2.Label3
        Case 5
            Form4.Show
            Form4.Label3.Caption = DataGrid1.Columns(0).Value
            Form4.Adodc1.Recordset.Filter = "ID like " + Form4.Label3
        Case 6
            Form5.Show
    End Select

End Sub

Private Sub Command2_Click()

    On Error Resume Next
    Adodc1.Recordset.Update
    Adodc1.Recordset.Requery

End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Adodc1.Recordset.Filter = "ESTATUS like 'Pendiente'"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Form2
    Unload Form4
    Unload Form5
    Unload Me
    
End Sub

Private Sub Option1_Click(Index As Integer)

    On Error Resume Next
    Select Case Index
        Case 1
            Adodc1.Recordset.Filter = "ESTATUS like 'Pendiente'"
        Case 2
            Adodc1.Recordset.Filter = "ESTATUS like 'En proceso'"
        Case 3
            Adodc1.Recordset.Filter = "ESTATUS like 'Resuelto'"
    End Select

End Sub
