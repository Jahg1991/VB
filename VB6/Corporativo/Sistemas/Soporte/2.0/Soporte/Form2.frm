VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisiòn de solicitudes"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   5400
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
      Connect         =   $"Form2.frx":72FA
      OLEDBString     =   $"Form2.frx":7387
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
      Left            =   2640
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informaciòn de la solicitud"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ESTATUS"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   18
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ATENDIDO"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   17
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "SOLICITANTE"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   14
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "DETALLES"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   4
         Left            =   1800
         TabIndex        =   13
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TIPO"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "HORA"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "FECHA"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ID"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Revisado por"
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   8
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Solicitante"
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   7
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Estatus"
         Height          =   375
         Index           =   6
         Left            =   960
         TabIndex        =   6
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Detalles"
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Hora"
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Id"
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command2_Click()
    
    Unload Me
    Form1.Show
    
End Sub

