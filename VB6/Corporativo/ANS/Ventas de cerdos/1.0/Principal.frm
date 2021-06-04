VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Venta de cerdos"
   ClientHeight    =   11580
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   18105
   ControlBox      =   0   'False
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Principal.frx":324A
   ScaleHeight     =   11580
   ScaleWidth      =   18105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FConsulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   17895
      Begin VB.PictureBox picImprimir 
         Height          =   255
         Left            =   2640
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   77
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "Buscar por..."
         Top             =   600
         Width           =   2295
      End
      Begin VB.Frame FClientes 
         Height          =   10095
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   17655
         Begin VB.TextBox Text21 
            DataField       =   "TEJABAN"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   13080
            TabIndex        =   115
            Top             =   10920
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox Text20 
            DataField       =   "TOTAL"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   11520
            TabIndex        =   114
            Top             =   10920
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ListBox List24 
            Height          =   8445
            Left            =   13080
            TabIndex        =   111
            Top             =   1440
            Width           =   3855
         End
         Begin VB.ListBox List23 
            Height          =   8445
            Left            =   11520
            TabIndex        =   110
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            DataField       =   "CLIENTE"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   7320
            TabIndex        =   22
            Top             =   10920
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox Text5 
            DataField       =   "PROMEDIO"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   6600
            TabIndex        =   21
            Top             =   10920
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text4 
            DataField       =   "KGS"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   5640
            TabIndex        =   20
            Top             =   10920
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text3 
            DataField       =   "NO"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   4560
            TabIndex        =   19
            Top             =   10920
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            DataField       =   "GRANJA"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   2160
            TabIndex        =   18
            Top             =   10920
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            DataField       =   "FECHA"
            DataSource      =   "Adodc1"
            Height          =   375
            Left            =   600
            TabIndex        =   17
            Top             =   10920
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   15480
            Top             =   10920
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\JAHG Software\Venta de cerdos\DB.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\JAHG Software\Venta de cerdos\DB.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "VC"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.ListBox List2 
            Height          =   8445
            Left            =   2160
            TabIndex        =   16
            Top             =   1440
            Width           =   2415
         End
         Begin VB.ListBox List6 
            Height          =   8445
            Left            =   7320
            TabIndex        =   15
            Top             =   1440
            Width           =   4215
         End
         Begin VB.ListBox List5 
            Height          =   8445
            Left            =   6600
            TabIndex        =   14
            Top             =   1440
            Width           =   735
         End
         Begin VB.ListBox List4 
            Height          =   8445
            Left            =   5640
            TabIndex        =   13
            Top             =   1440
            Width           =   975
         End
         Begin VB.ListBox List3 
            Height          =   8445
            Left            =   4560
            TabIndex        =   12
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ListBox List1 
            Height          =   8445
            Left            =   600
            TabIndex        =   11
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox BClientes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   240
            Width           =   6975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Tejaban"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   13080
            TabIndex        =   113
            Top             =   1080
            Width           =   3855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   11520
            TabIndex        =   112
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   7320
            TabIndex        =   10
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Prom."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   6600
            TabIndex        =   9
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Kgs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5640
            TabIndex        =   8
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4560
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Granja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   6
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   5
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FTodo 
         Height          =   10095
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   17655
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "Principal.frx":41669
            Height          =   9615
            Left            =   120
            Negotiate       =   -1  'True
            TabIndex        =   38
            Top             =   360
            Width           =   17415
            _ExtentX        =   30718
            _ExtentY        =   16960
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Resumen Venta de cerdos"
            ColumnCount     =   14
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "GRANJA"
               Caption         =   "GRANJA"
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
               DataField       =   "NO"
               Caption         =   "NO"
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
               DataField       =   "KGS"
               Caption         =   "KGS"
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
               DataField       =   "PROMEDIO"
               Caption         =   "PROMEDIO"
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
               DataField       =   "$KG"
               Caption         =   "$KG"
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
               DataField       =   "SUBTOTAL"
               Caption         =   "SUBTOTAL"
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
               DataField       =   "GUIAS"
               Caption         =   "GUIAS"
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
               DataField       =   "COMISIONES"
               Caption         =   "COMISIONES"
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
            BeginProperty Column09 
               DataField       =   "TOTAL"
               Caption         =   "TOTAL"
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
            BeginProperty Column10 
               DataField       =   "CLIENTE"
               Caption         =   "CLIENTE"
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
            BeginProperty Column11 
               DataField       =   "TEJABAN"
               Caption         =   "TEJABAN"
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
            BeginProperty Column12 
               DataField       =   "MORTALIDAD"
               Caption         =   "MORTALIDAD"
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
            BeginProperty Column13 
               DataField       =   "OBSERVACION"
               Caption         =   "OBSERVACION"
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
                  Locked          =   -1  'True
                  ColumnWidth     =   1500,095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1500,095
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column07 
                  Locked          =   -1  'True
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column10 
                  Locked          =   -1  'True
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  ColumnWidth     =   1124,787
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  ColumnWidth     =   1544,882
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FFecha 
         Height          =   10095
         Left            =   120
         TabIndex        =   82
         Top             =   1200
         Visible         =   0   'False
         Width           =   17655
         Begin VB.ListBox List22 
            Height          =   8445
            Left            =   13080
            TabIndex        =   108
            Top             =   1440
            Width           =   3855
         End
         Begin VB.ListBox List21 
            Height          =   8445
            Left            =   11520
            TabIndex        =   107
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Buscar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8520
            TabIndex        =   92
            Top             =   240
            Width           =   1335
         End
         Begin VB.ListBox List13 
            Height          =   8445
            Left            =   600
            TabIndex        =   90
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ListBox List14 
            Height          =   8445
            Left            =   2160
            TabIndex        =   89
            Top             =   1440
            Width           =   2415
         End
         Begin VB.ListBox List15 
            Height          =   8445
            Left            =   4560
            TabIndex        =   88
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ListBox List16 
            Height          =   8445
            Left            =   5640
            TabIndex        =   87
            Top             =   1440
            Width           =   975
         End
         Begin VB.ListBox List17 
            Height          =   8445
            Left            =   6600
            TabIndex        =   86
            Top             =   1440
            Width           =   735
         End
         Begin VB.ListBox List18 
            Height          =   8445
            Left            =   7320
            TabIndex        =   85
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox Fecha_inicial 
            Height          =   375
            Left            =   2400
            TabIndex        =   84
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox Fecha_final 
            Height          =   375
            Left            =   6240
            TabIndex        =   83
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   2400
            TabIndex        =   91
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   97452033
            CurrentDate     =   41494
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   6240
            TabIndex        =   93
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   97452033
            CurrentDate     =   41494
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Tejaban"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   13080
            TabIndex        =   109
            Top             =   1080
            Width           =   3855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   11520
            TabIndex        =   106
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde la fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   101
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta la fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            TabIndex        =   100
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   600
            TabIndex        =   99
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Granja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   2160
            TabIndex        =   98
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   4560
            TabIndex        =   97
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Kgs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   5760
            TabIndex        =   96
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Prom"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   6600
            TabIndex        =   95
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C000&
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   7320
            TabIndex        =   94
            Top             =   1080
            Width           =   4215
         End
      End
      Begin VB.Image Image3 
         Height          =   10080
         Left            =   120
         Picture         =   "Principal.frx":4167E
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   17595
      End
   End
   Begin VB.Frame FIngreso 
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   120
      TabIndex        =   78
      Top             =   120
      Width           =   17895
      Begin VB.CommandButton Command9 
         BackColor       =   &H00808000&
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   10080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   80
         Top             =   10080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   79
         Text            =   "Usuario"
         Top             =   10080
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   11055
         Left            =   0
         Picture         =   "Principal.frx":454DF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   17895
      End
   End
   Begin VB.Frame FGranja 
      Height          =   10095
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   17655
      Begin VB.ListBox List20 
         Height          =   8445
         Left            =   13080
         TabIndex        =   105
         Top             =   1440
         Width           =   3855
      End
      Begin VB.ListBox List19 
         Height          =   8445
         Left            =   11520
         TabIndex        =   103
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   600
         TabIndex        =   36
         Text            =   "Granja..."
         Top             =   360
         Width           =   2415
      End
      Begin VB.ListBox List7 
         Height          =   8445
         Left            =   600
         TabIndex        =   29
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ListBox List9 
         Height          =   8445
         Left            =   4560
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List10 
         Height          =   8445
         Left            =   5640
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox List11 
         Height          =   8445
         Left            =   6600
         TabIndex        =   26
         Top             =   1440
         Width           =   735
      End
      Begin VB.ListBox List12 
         Height          =   8445
         Left            =   7320
         TabIndex        =   25
         Top             =   1440
         Width           =   4215
      End
      Begin VB.ListBox List8 
         Height          =   8445
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Tejaban"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   13080
         TabIndex        =   104
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "$ Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   11520
         TabIndex        =   102
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   600
         TabIndex        =   35
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Granja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   2160
         TabIndex        =   34
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4560
         TabIndex        =   33
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Kgs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   5640
         TabIndex        =   32
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Prom."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   6600
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   7320
         TabIndex        =   30
         Top             =   1080
         Width           =   4215
      End
   End
   Begin VB.Frame FRegistro 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   17895
      Begin VB.Frame FNew 
         Height          =   10815
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   17655
         Begin VB.CommandButton Command8 
            Caption         =   "Último registro"
            Height          =   375
            Left            =   10200
            TabIndex        =   76
            Top             =   10320
            Width           =   1935
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Siguiente registro"
            Height          =   375
            Left            =   8160
            TabIndex        =   75
            Top             =   10320
            Width           =   1935
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Registro anterior"
            Height          =   375
            Left            =   6120
            TabIndex        =   74
            Top             =   10320
            Width           =   1935
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Primer registro"
            Height          =   375
            Left            =   4080
            TabIndex        =   73
            Top             =   10320
            Width           =   1935
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Eliminar"
            Height          =   375
            Left            =   15600
            TabIndex        =   72
            Top             =   10320
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Guardar cambios"
            Height          =   375
            Left            =   13560
            TabIndex        =   71
            Top             =   10320
            Width           =   1935
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "Principal.frx":4EECE
            Height          =   9495
            Left            =   4080
            TabIndex        =   70
            Top             =   720
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   16748
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   14
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   14
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "GRANJA"
               Caption         =   "GRANJA"
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
               DataField       =   "NO"
               Caption         =   "NO"
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
               DataField       =   "KGS"
               Caption         =   "KGS"
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
               DataField       =   "PROMEDIO"
               Caption         =   "PROM"
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
               DataField       =   "$KG"
               Caption         =   "$KG"
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
               DataField       =   "SUBTOTAL"
               Caption         =   "SUBTOTAL"
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
               DataField       =   "GUIAS"
               Caption         =   "GUIAS"
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
               DataField       =   "COMISIONES"
               Caption         =   "COM."
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
            BeginProperty Column09 
               DataField       =   "TOTAL"
               Caption         =   "TOTAL"
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
            BeginProperty Column10 
               DataField       =   "CLIENTE"
               Caption         =   "CLIENTE"
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
            BeginProperty Column11 
               DataField       =   "TEJABAN"
               Caption         =   "TEJABAN"
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
            BeginProperty Column12 
               DataField       =   "MORTALIDAD"
               Caption         =   "MORT."
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
            BeginProperty Column13 
               DataField       =   "OBSERVACION"
               Caption         =   "OBSERVACION"
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
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   615,118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   689,953
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   854,929
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   585,071
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   675,213
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   585,071
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   884,976
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1604,976
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   629,858
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   1335,118
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   1440
            TabIndex        =   69
            Top             =   7680
            Width           =   1935
         End
         Begin VB.TextBox Text18 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   68
            Text            =   "-"
            Top             =   6960
            Width           =   2295
         End
         Begin VB.TextBox Text17 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   66
            Text            =   "0"
            Top             =   6480
            Width           =   2295
         End
         Begin VB.TextBox Text16 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   64
            Text            =   "-"
            Top             =   6000
            Width           =   2295
         End
         Begin VB.TextBox Text15 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   62
            Text            =   "-"
            Top             =   5520
            Width           =   2295
         End
         Begin VB.TextBox Text14 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "0"
            Top             =   5040
            Width           =   2295
         End
         Begin VB.TextBox Text13 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   58
            Text            =   "0"
            Top             =   4560
            Width           =   2295
         End
         Begin VB.TextBox Text12 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   56
            Text            =   "0"
            Top             =   4080
            Width           =   2295
         End
         Begin VB.TextBox Text11 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "0"
            Top             =   3600
            Width           =   2295
         End
         Begin VB.TextBox Text10 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   52
            Text            =   "0"
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Text8 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   50
            Text            =   "0"
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "0"
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox Text7 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   46
            Text            =   "0"
            Top             =   1680
            Width           =   2295
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1680
            TabIndex        =   43
            Text            =   "-"
            Top             =   1200
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1680
            TabIndex        =   41
            Top             =   720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            Format          =   97452033
            CurrentDate     =   41494
         End
         Begin VB.Image Image2 
            Height          =   2175
            Left            =   120
            Picture         =   "Principal.frx":4EEE3
            Stretch         =   -1  'True
            Top             =   8520
            Width           =   3855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   0
            TabIndex        =   67
            Top             =   6960
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mortalidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   0
            TabIndex        =   65
            Top             =   6480
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tejaban"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   0
            TabIndex        =   63
            Top             =   6000
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   0
            TabIndex        =   61
            Top             =   5520
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   -120
            TabIndex        =   59
            Top             =   5040
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comisiones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   57
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Guias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   55
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   53
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$ Kg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   51
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Promedio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   49
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kilogramos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   47
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   45
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Granja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   44
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   42
            Top             =   720
            Width           =   735
         End
      End
   End
   Begin VB.Menu Registro 
      Caption         =   "Registro"
   End
   Begin VB.Menu Consulta 
      Caption         =   "Consulta"
   End
   Begin VB.Menu Imprimir 
      Caption         =   "Imprimir"
   End
   Begin VB.Menu EAE 
      Caption         =   "Exportar a excel"
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Function Exportar_ADO_Excel(sPathDB As String, Sql As String, sOutputPathXLS As String) As Boolean
    
    On Error GoTo ErrSub
    
    Dim cn1          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    
    Me.Enabled = False
    
   ' -- Abrir la base
    cn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"
        
    ' -- Abrir el Recordset pasándole la cadena sql
    rec.Open Sql, cn1
    
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
    
    Excel.Visible = True: Excel.UserControl = True
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
    Next
    
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rec
    Else

        arrData = rec.GetRows

        iRec = UBound(arrData, 2) + 1
        
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1

                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))

                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
            
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
    End If

    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit

    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
    cn1.Close
    
    Set rec = Nothing
    Set cn1 = Nothing
    ' -- guardar el libro
    Libro.saveAs sOutputPathXLS
    Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    Excel.quit
    Set Excel = Nothing
    
    Exportar_ADO_Excel = True
    Me.Enabled = True
    Exit Function
ErrSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel = False
    Me.Enabled = True
End Function

Private Function GetData(vValue As Variant) As Variant
    Dim x As Long, y As Long, xMax As Long, yMax As Long, T As Variant
    
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
    
    ReDim T(xMax, yMax)
    For x = 0 To xMax
        For y = 0 To yMax
            T(x, y) = vValue(y, x)
        Next y
    Next x
    
    GetData = T
End Function

Private Sub BClientes_Change()

On Error Resume Next
With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List1.Clear
Do While Not .EOF
List1.AddItem rs("FECHA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List2.Clear
Do While Not .EOF
List2.AddItem rs("GRANJA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List3.Clear
Do While Not .EOF
List3.AddItem rs("NO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List4.Clear
Do While Not .EOF
List4.AddItem rs("KGS")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List5.Clear
Do While Not .EOF
List5.AddItem rs("PROMEDIO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List6.Clear
Do While Not .EOF
List6.AddItem rs("CLIENTE")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List23.Clear
Do While Not .EOF
List23.AddItem rs("TOTAL")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "CLIENTE LIKE '*" & BClientes & "*'"
End If
List24.Clear
Do While Not .EOF
List24.AddItem rs("TEJABAN")
.MoveNext
Loop
End With

End Sub

Private Sub BClientes_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
List1.SetFocus

End If

End Sub

Private Sub Combo1_Click()

On Error Resume Next
Select Case Combo1.Text

Case "Fecha"
On Error Resume Next
FClientes.Visible = False
FGranja.Visible = False
FTodo.Visible = False
FFecha.Visible = True
Combo2.Text = "Granja..."
DTPicker1.Value = Date
DTPicker2.Value = Date
BClientes = ""

Case "Cliente"
On Error Resume Next
FClientes.Visible = True
FGranja.Visible = False
FTodo.Visible = False
FFecha.Visible = False
Combo2.Text = "Granja..."
DTPicker1.Value = Date
DTPicker2.Value = Date
BClientes = ""

Case "Granja"
On Error Resume Next
FClientes.Visible = False
FGranja.Visible = True
FTodo.Visible = False
FFecha.Visible = False
Combo2.Text = "Granja..."
DTPicker1.Value = Date
DTPicker2.Value = Date
BClientes = ""

Case "Todo"
On Error Resume Next
FClientes.Visible = False
FGranja.Visible = False
FTodo.Visible = True
FFecha.Visible = False
Combo2.Text = "Granja..."
DTPicker1.Value = Date
DTPicker2.Value = Date
BClientes = ""

End Select

End Sub

Private Sub Combo2_Click()
On Error Resume Next
With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List7.Clear
Do While Not .EOF
List7.AddItem rs("FECHA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List8.Clear
Do While Not .EOF
List8.AddItem rs("GRANJA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List9.Clear
Do While Not .EOF
List9.AddItem rs("NO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List10.Clear
Do While Not .EOF
List10.AddItem rs("KGS")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List11.Clear
Do While Not .EOF
List11.AddItem rs("PROMEDIO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List12.Clear
Do While Not .EOF
List12.AddItem rs("CLIENTE")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List19.Clear
Do While Not .EOF
List19.AddItem rs("TOTAL")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "GRANJA LIKE '*" & Combo2 & "*'"
End If
List20.Clear
Do While Not .EOF
List20.AddItem rs("TEJABAN")
.MoveNext
Loop
End With

End Sub

Private Sub Combo3_Click()

On Error Resume Next
Select Case Combo3.Text

Case "Consulta"
On Error Resume Next
FIngreso.Visible = False
Registro.Enabled = False
Consulta.Enabled = True
Imprimir.Enabled = True
EAE.Enabled = True

Case "Registro"

On Error Resume Next
Text19.Visible = True
Command9.Visible = True
Text19.SetFocus

End Select

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text7.SetFocus
End If

End Sub

Private Sub Command1_Click()

On Error Resume Next
List13.Clear
List14.Clear
List15.Clear
List16.Clear
List17.Clear
List18.Clear
List21.Clear
List22.Clear

Fecha_inicial.Text = DTPicker1.Value
Fecha_final.Text = DTPicker2.Value

Command1.Enabled = False

End Sub

Private Sub Command2_Click()

On Error Resume Next
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("FECHA") = DTPicker3.Value
Adodc1.Recordset.Fields("GRANJA") = Combo4.Text
Adodc1.Recordset.Fields("NO") = Text7.Text
Adodc1.Recordset.Fields("KGS") = Text8.Text
Adodc1.Recordset.Fields("PROMEDIO") = Text9.Text
Adodc1.Recordset.Fields("$KG") = Text10.Text
Adodc1.Recordset.Fields("SUBTOTAL") = Text11.Text
Adodc1.Recordset.Fields("GUIAS") = Text12.Text
Adodc1.Recordset.Fields("COMISIONES") = Text13.Text
Adodc1.Recordset.Fields("TOTAL") = Text14.Text
Adodc1.Recordset.Fields("CLIENTE") = Text15.Text
Adodc1.Recordset.Fields("TEJABAN") = Text16.Text
Adodc1.Recordset.Fields("MORTALIDAD") = Text17.Text
Adodc1.Recordset.Fields("OBSERVACION") = Text18.Text
Adodc1.Recordset.Update
MsgBox ("Venta guardada con exito")
Text7.Text = "0"
Text8.Text = "0"
Text9.Text = "0"
Text10.Text = "0"
Text11.Text = "0"
Text12.Text = "0"
Text13.Text = "0"
Text14.Text = "0"
Text15.Text = "-"
Text16.Text = "-"
Text17.Text = "0"
Text18.Text = "-"
Combo4.Text = "-"
DTPicker3.Value = Date
DTPicker3.SetFocus

End Sub

Private Sub Command3_Click()

On Error Resume Next
Adodc1.Recordset.Update
MsgBox ("Cambios guardados exitosamente")

End Sub

Private Sub Command4_Click()

On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Recordset.Update
MsgBox ("Registro eliminado")

End Sub

Private Sub Command5_Click()

On Error Resume Next
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command6_Click()

On Error Resume Next
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command7_Click()

On Error Resume Next
Adodc1.Recordset.MoveNext

End Sub

Private Sub Command8_Click()

On Error Resume Next
Adodc1.Recordset.MoveLast

End Sub

Private Sub Command9_Click()

On Error Resume Next
If Text19.Text = "vencerns" Then
FIngreso.Visible = False
Text19.Text = ""
Consulta.Enabled = True
Imprimir.Enabled = True
EAE.Enabled = True
Registro.Enabled = True
Else
If Text19.Text < vencerns > "" Then
MsgBox "Contraseña incorrecta"
Text19.Text = ""
Text19.SetFocus
End If
End If

End Sub

Private Sub Consulta_Click()

On Error Resume Next
FConsulta.Visible = True
FRegistro.Visible = False

End Sub

Private Sub CP_Click()

On Error Resume Next
Call Configuarar_Pagina(Me.hWnd)

End Sub

Private Sub DTPicker1_CloseUp()

On Error Resume Next
Command1.Enabled = True

End Sub

Private Sub DTPicker2_CloseUp()

On Error Resume Next
Command1.Enabled = True

End Sub

Private Sub DTPicker3_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Combo4.SetFocus
End If

End Sub

Private Sub EAE_Click()

    Dim sPathDB        As String
    Dim Consulta    As String

    ' -- Path de la base de datos
    sPathDB = "C:\JAHG Software\Venta de cerdos\DB.MDB"

    ' -- Cadena Sql
    Consulta = "Select * From VC"

    ' -- Enviar el Path de la base de datos y la consulta sql
    If Exportar_ADO_Excel(sPathDB, Consulta, "C:\JAHG Software\Venta de cerdos\Reportes\Reporte.xLS") Then
       MsgBox "Ok", vbInformation
    End If


End Sub

Private Sub Fecha_final_Change()

On Error Resume Next

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List13.Clear
Do While Not .EOF
List13.AddItem rs("FECHA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List14.Clear
Do While Not .EOF
List14.AddItem rs("GRANJA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List15.Clear
Do While Not .EOF
List15.AddItem rs("NO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List16.Clear
Do While Not .EOF
List16.AddItem rs("KGS")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List17.Clear
Do While Not .EOF
List17.AddItem rs("PROMEDIO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List18.Clear
Do While Not .EOF
List18.AddItem rs("CLIENTE")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List21.Clear
Do While Not .EOF
List21.AddItem rs("TOTAL")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List22.Clear
Do While Not .EOF
List22.AddItem rs("TEJABAN")
.MoveNext
Loop
End With

End Sub

Private Sub Fecha_inicial_Change()
On Error Resume Next

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List13.Clear
Do While Not .EOF
List13.AddItem rs("FECHA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List14.Clear
Do While Not .EOF
List14.AddItem rs("GRANJA")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List15.Clear
Do While Not .EOF
List15.AddItem rs("NO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List16.Clear
Do While Not .EOF
List16.AddItem rs("KGS")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List17.Clear
Do While Not .EOF
List17.AddItem rs("PROMEDIO")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List18.Clear
Do While Not .EOF
List18.AddItem rs("CLIENTE")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List21.Clear
Do While Not .EOF
List21.AddItem rs("TOTAL")
.MoveNext
Loop
End With

With rs
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
End If
List22.Clear
Do While Not .EOF
List22.AddItem rs("TEJABAN")
.MoveNext
Loop
End With

End Sub

Private Sub Form_Load()

On Error Resume Next
Combo1.AddItem "Fecha"
Combo1.AddItem "Cliente"
Combo1.AddItem "Granja"
Combo1.AddItem "Todo"

Combo2.AddItem "Terrero"
Combo2.AddItem "Isabel"
Combo2.AddItem "Laja"
Combo2.AddItem "Cuna"
Combo2.AddItem "Sapo"
Combo2.AddItem "Moro"
Combo2.AddItem "Loma"

Combo3.AddItem "Consulta"
Combo3.AddItem "Registro"

Combo4.AddItem "Terrero"
Combo4.AddItem "Isabel"
Combo4.AddItem "Laja"
Combo4.AddItem "Cuna"
Combo4.AddItem "Sapo"
Combo4.AddItem "Moro"
Combo4.AddItem "Loma"

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\DB.mdb;Persist Security Info=False"
Set rs = cn.Execute("SELECT * FROM VC")

DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date

Registro.Enabled = False
Consulta.Enabled = False
Imprimir.Enabled = False
EAE.Enabled = False

End Sub

Private Sub Imprimir_Click()

On Error Resume Next

Printer.Orientation = vbPRORLandscape
Printer.PaperSize = vbPRPSLetter 'Tipo de Papel

picImprimir.Picture = CaptureClient(Me)
Printer.PaintPicture picImprimir.Picture, 0, 0, Printer.ScaleWidth, (Me.ScaleHeight * Printer.ScaleWidth) / Me.ScaleWidth, , , Me.ScaleWidth, Me.ScaleHeight
Printer.EndDoc

End Sub

Private Sub List1_Click()

On Error Resume Next
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
List6.ListIndex = List1.ListIndex
List23.ListIndex = List1.ListIndex
List24.ListIndex = List1.ListIndex

End Sub

Private Sub List10_Click()

On Error Resume Next
List8.ListIndex = List10.ListIndex
List9.ListIndex = List10.ListIndex
List7.ListIndex = List10.ListIndex
List11.ListIndex = List10.ListIndex
List12.ListIndex = List10.ListIndex
List19.ListIndex = List10.ListIndex
List20.ListIndex = List10.ListIndex

End Sub

Private Sub List11_Click()

On Error Resume Next
List8.ListIndex = List11.ListIndex
List9.ListIndex = List11.ListIndex
List10.ListIndex = List11.ListIndex
List7.ListIndex = List11.ListIndex
List12.ListIndex = List11.ListIndex
List19.ListIndex = List11.ListIndex
List20.ListIndex = List11.ListIndex

End Sub

Private Sub List12_Click()

On Error Resume Next
List8.ListIndex = List12.ListIndex
List9.ListIndex = List12.ListIndex
List10.ListIndex = List12.ListIndex
List11.ListIndex = List12.ListIndex
List7.ListIndex = List12.ListIndex
List19.ListIndex = List12.ListIndex
List20.ListIndex = List12.ListIndex

End Sub

Private Sub List13_Click()

On Error Resume Next
List14.ListIndex = List13.ListIndex
List15.ListIndex = List13.ListIndex
List16.ListIndex = List13.ListIndex
List17.ListIndex = List13.ListIndex
List18.ListIndex = List13.ListIndex
List21.ListIndex = List13.ListIndex
List22.ListIndex = List13.ListIndex

End Sub

Private Sub List14_Click()

On Error Resume Next
List13.ListIndex = List14.ListIndex
List15.ListIndex = List14.ListIndex
List16.ListIndex = List14.ListIndex
List17.ListIndex = List14.ListIndex
List18.ListIndex = List14.ListIndex
List21.ListIndex = List14.ListIndex
List22.ListIndex = List14.ListIndex


End Sub

Private Sub List15_Click()

On Error Resume Next
List14.ListIndex = List15.ListIndex
List13.ListIndex = List15.ListIndex
List16.ListIndex = List15.ListIndex
List17.ListIndex = List15.ListIndex
List18.ListIndex = List15.ListIndex
List21.ListIndex = List15.ListIndex
List22.ListIndex = List15.ListIndex

End Sub

Private Sub List16_Click()

On Error Resume Next
List14.ListIndex = List16.ListIndex
List15.ListIndex = List16.ListIndex
List13.ListIndex = List16.ListIndex
List17.ListIndex = List16.ListIndex
List18.ListIndex = List16.ListIndex
List21.ListIndex = List16.ListIndex
List22.ListIndex = List16.ListIndex

End Sub

Private Sub List17_Click()

On Error Resume Next
List14.ListIndex = List17.ListIndex
List15.ListIndex = List17.ListIndex
List16.ListIndex = List17.ListIndex
List13.ListIndex = List17.ListIndex
List18.ListIndex = List17.ListIndex
List21.ListIndex = List17.ListIndex
List22.ListIndex = List17.ListIndex

End Sub

Private Sub List18_Click()

On Error Resume Next
List14.ListIndex = List18.ListIndex
List15.ListIndex = List18.ListIndex
List16.ListIndex = List18.ListIndex
List17.ListIndex = List18.ListIndex
List13.ListIndex = List18.ListIndex
List21.ListIndex = List18.ListIndex
List22.ListIndex = List18.ListIndex

End Sub

Private Sub List19_Click()

On Error Resume Next
List8.ListIndex = List19.ListIndex
List9.ListIndex = List19.ListIndex
List10.ListIndex = List19.ListIndex
List11.ListIndex = List19.ListIndex
List7.ListIndex = List19.ListIndex
List12.ListIndex = List19.ListIndex
List20.ListIndex = List19.ListIndex

End Sub

Private Sub List2_Click()

On Error Resume Next
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List5.ListIndex = List2.ListIndex
List6.ListIndex = List2.ListIndex
List23.ListIndex = List2.ListIndex
List24.ListIndex = List2.ListIndex

End Sub

Private Sub List20_Click()

On Error Resume Next
List8.ListIndex = List20.ListIndex
List9.ListIndex = List20.ListIndex
List10.ListIndex = List20.ListIndex
List11.ListIndex = List20.ListIndex
List7.ListIndex = List20.ListIndex
List12.ListIndex = List20.ListIndex
List19.ListIndex = List20.ListIndex

End Sub

Private Sub List21_Click()

On Error Resume Next
List14.ListIndex = List21.ListIndex
List15.ListIndex = List21.ListIndex
List16.ListIndex = List21.ListIndex
List17.ListIndex = List21.ListIndex
List13.ListIndex = List21.ListIndex
List18.ListIndex = List21.ListIndex
List22.ListIndex = List21.ListIndex

End Sub

Private Sub List22_Click()

On Error Resume Next
List14.ListIndex = List22.ListIndex
List15.ListIndex = List22.ListIndex
List16.ListIndex = List22.ListIndex
List17.ListIndex = List22.ListIndex
List13.ListIndex = List22.ListIndex
List18.ListIndex = List22.ListIndex
List21.ListIndex = List22.ListIndex

End Sub

Private Sub List23_Click()

On Error Resume Next
List1.ListIndex = List23.ListIndex
List3.ListIndex = List23.ListIndex
List4.ListIndex = List23.ListIndex
List5.ListIndex = List23.ListIndex
List2.ListIndex = List23.ListIndex
List6.ListIndex = List23.ListIndex
List24.ListIndex = List23.ListIndex

End Sub

Private Sub List24_Click()

On Error Resume Next
List1.ListIndex = List24.ListIndex
List3.ListIndex = List24.ListIndex
List4.ListIndex = List24.ListIndex
List5.ListIndex = List24.ListIndex
List2.ListIndex = List24.ListIndex
List6.ListIndex = List24.ListIndex
List23.ListIndex = List24.ListIndex

End Sub

Private Sub List3_Click()

On Error Resume Next
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
List6.ListIndex = List3.ListIndex
List23.ListIndex = List3.ListIndex
List24.ListIndex = List3.ListIndex

End Sub

Private Sub List4_Click()

On Error Resume Next
List1.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
List6.ListIndex = List4.ListIndex
List23.ListIndex = List4.ListIndex
List24.ListIndex = List4.ListIndex

End Sub

Private Sub List5_Click()

On Error Resume Next
List1.ListIndex = List5.ListIndex
List3.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
List2.ListIndex = List5.ListIndex
List6.ListIndex = List5.ListIndex
List23.ListIndex = List5.ListIndex
List24.ListIndex = List5.ListIndex

End Sub

Private Sub List6_Click()

On Error Resume Next
List1.ListIndex = List6.ListIndex
List3.ListIndex = List6.ListIndex
List4.ListIndex = List6.ListIndex
List5.ListIndex = List6.ListIndex
List2.ListIndex = List6.ListIndex
List23.ListIndex = List6.ListIndex
List24.ListIndex = List6.ListIndex

End Sub

Private Sub List7_Click()

On Error Resume Next
List8.ListIndex = List7.ListIndex
List9.ListIndex = List7.ListIndex
List10.ListIndex = List7.ListIndex
List11.ListIndex = List7.ListIndex
List12.ListIndex = List7.ListIndex
List19.ListIndex = List7.ListIndex
List20.ListIndex = List7.ListIndex

End Sub

Private Sub List8_Click()

On Error Resume Next
List7.ListIndex = List8.ListIndex
List9.ListIndex = List8.ListIndex
List10.ListIndex = List8.ListIndex
List11.ListIndex = List8.ListIndex
List12.ListIndex = List8.ListIndex
List19.ListIndex = List8.ListIndex
List20.ListIndex = List8.ListIndex

End Sub

Private Sub List9_Click()

On Error Resume Next
List8.ListIndex = List9.ListIndex
List7.ListIndex = List9.ListIndex
List10.ListIndex = List9.ListIndex
List11.ListIndex = List9.ListIndex
List12.ListIndex = List9.ListIndex
List19.ListIndex = List9.ListIndex
List20.ListIndex = List9.ListIndex

End Sub

Private Sub Print_Click()

On Error Resume Next
Call Show_Printer(Me)

End Sub

Private Sub picImprimir_Click()

End Sub

Private Sub Registro_Click()

On Error Resume Next
FConsulta.Visible = False
FRegistro.Visible = True

End Sub

Private Sub Salir_Click()

On Error Resume Next
Unload Form1

End Sub

Private Sub Text10_Change()

On Error Resume Next
Text11 = Text8 * Text10

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text11.SetFocus
End If

End Sub

Private Sub Text11_Change()

On Error Resume Next
Text14.Text = Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text)

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text12.SetFocus
End If

End Sub

Private Sub Text12_Change()

On Error Resume Next
Text14.Text = Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text13.SetFocus
End If

End Sub

Private Sub Text13_Change()

On Error Resume Next
Text14.Text = Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text)

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text14.SetFocus
End If

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text15.SetFocus
End If

End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text16.SetFocus
End If

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text17.SetFocus
End If

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text18.SetFocus
End If

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Command2.SetFocus
End If

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Command9.SetFocus
End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text8.SetFocus
End If

End Sub

Private Sub Text8_Change()

On Error Resume Next
Text9 = Text8 / Text7
Text11 = Text8 * Text10

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text9.SetFocus
End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 13 Then
Text10.SetFocus
End If

End Sub
