VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estatus de la solicitud"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form7.frx":164A
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   5640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1296
      _Version        =   393216
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "TIPO"
         Caption         =   "TIPO"
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
         DataField       =   "DETALLES"
         Caption         =   "DETALLES"
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
         DataField       =   "IMAGEN"
         Caption         =   "IMAGEN"
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
         DataField       =   "FECHA"
         Caption         =   "FECHA"
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
         DataField       =   "HORA"
         Caption         =   "HORA"
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
      BeginProperty Column06 
         DataField       =   "ESTATUS"
         Caption         =   "ESTATUS"
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
      BeginProperty Column07 
         DataField       =   "SOLICITANTE"
         Caption         =   "SOLICITANTE"
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
      BeginProperty Column08 
         DataField       =   "ATENDIDO"
         Caption         =   "ATENDIDO"
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
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2085,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   4800
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
      CommandType     =   2
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
      Connect         =   $"Form7.frx":165F
      OLEDBString     =   $"Form7.frx":16E6
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SOLICITUDES"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cambiar estatus"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ID"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Id"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Revisado por"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Estatus"
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   1
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click(Index As Integer)

    Select Case Index
        Case 0
            DataGrid1.Columns(8).Text = Combo1(0).Text
            Adodc1.Recordset.Update
        Case 1
            DataGrid1.Columns(6).Text = Combo1(1).Text
            Adodc1.Recordset.Update
    End Select
        
End Sub

Private Sub Command2_Click()
    
    Unload Form1
    Unload Me
    Form1.Show
    
End Sub

Private Sub Form_Load()

    Combo1(0).AddItem ("Alfredo Hernàndez")
    Combo1(0).AddItem ("Berenice Hernàndez")
    Combo1(0).AddItem ("Eduardo Carreòn")
    Combo1(0).AddItem ("Guadalupe Flores")
    Combo1(0).AddItem ("Guadalupe Villanueva")
    Combo1(0).AddItem ("Samuel Padilla")

    Combo1(1).AddItem ("Pendiente")
    Combo1(1).AddItem ("En proceso")
    Combo1(1).AddItem ("Resuelto")
            
End Sub
