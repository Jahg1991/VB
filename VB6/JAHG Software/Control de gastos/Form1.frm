VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gastos"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
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
   ScaleHeight     =   8235
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1538
      TabIndex        =   34
      Top             =   7680
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   167903233
      CurrentDate     =   43784
      MaxDate         =   72686
      MinDate         =   43466
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ingresos"
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   5760
      Width           =   5175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   1680
         TabIndex        =   33
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   1680
         TabIndex        =   32
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inbursa"
         Height          =   375
         Index           =   14
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yisel"
         Height          =   375
         Index           =   13
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Alfredo"
         Height          =   375
         Index           =   12
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gastos"
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1680
         TabIndex        =   30
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1680
         TabIndex        =   29
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   28
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   27
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1680
         TabIndex        =   26
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   25
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   24
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   23
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   22
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Viajes"
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salidas"
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Renta"
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Otros"
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inbursa"
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandado"
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Yisel"
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Alfredo"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hogar"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Carro"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Casa"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   2040
      TabIndex        =   19
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Presupuesto $"
      Height          =   375
      Index           =   15
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu Ingresos 
         Caption         =   "Ingresos"
         Begin VB.Menu ISemana 
            Caption         =   "Semana"
         End
         Begin VB.Menu IMes 
            Caption         =   "Mes"
         End
      End
      Begin VB.Menu Gastos 
         Caption         =   "Gastos"
         Begin VB.Menu GSemana 
            Caption         =   "Semana"
         End
         Begin VB.Menu GMes 
            Caption         =   "Mes"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    guardar
    
    InitForm

End Sub

Private Sub GMes_Click()

    TipoReporte = 4
    CExportarExcel
    
End Sub

Private Sub GSemana_Click()

    TipoReporte = 3
    CExportarExcel
    
End Sub

Private Sub IMes_Click()

    TipoReporte = 2
    CExportarExcel
    
End Sub

Private Sub ISemana_Click()

    TipoReporte = 1
    CExportarExcel
    
End Sub
