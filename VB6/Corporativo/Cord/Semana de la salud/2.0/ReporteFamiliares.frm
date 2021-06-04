VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form ReporteFamiliares 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte familiares asistentes por trabajador"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   9015
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReporteFamiliares.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "Último"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6000
         TabIndex        =   15
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   14
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   13
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Primero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de nacimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   6480
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Parentesco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   10
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
      DefColWidth     =   267
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre Completo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu AArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "ReporteFamiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
    With RSTrabajador
        Select Case Index
            Case 0
                .MoveFirst
            Case 1
                .MovePrevious
            Case 2
                .MoveNext
            Case 3
                .MoveLast
        End Select
    End With
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RSTrabajador
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT Where TRAB_E = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsFamiliar
        If .State = 1 Then .Close
            .Open "Select ID_EMP, NOMBRE, PARENT, FE_NAC from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set DataGrid1.DataSource = RSTrabajador
    Set DataGrid2.DataSource = RsFamiliar
    Set Label1(1).DataSource = RSTrabajador
    Label1(1).DataField = ("ID_AST")
    Set Label1(5).DataSource = RSTrabajador
    Label1(5).DataField = ("NOMBRE")
    DataGrid2.Columns(1).Width = 4095
    DataGrid2.Columns(2).Width = 1935
    DataGrid2.Columns(3).Width = 1935
    DataGrid2.Columns(0).Visible = False
End Sub
Private Sub Imprimir_Click()
    On Error Resume Next
    Set DataReport2.DataSource = RsFamiliar
    DRID_AST = ReporteFamiliares.Label1(1).Caption
    DRNOMBRE = ReporteFamiliares.Label1(5).Caption
    DataReport2.Sections("Sección2").Controls("Etiqueta2").Caption = DRID_AST
    DataReport2.Sections("Sección2").Controls("Etiqueta3").Caption = DRNOMBRE
    DataReport2.Show
End Sub
Private Sub Label1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1
            With RsFamiliar
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "ID_EMP LIKE '*" & Label1(1) & "*'"
                    DataGrid2.Columns(1).Width = 4095
                    DataGrid2.Columns(2).Width = 1935
                    DataGrid2.Columns(3).Width = 1935
                    DataGrid2.Columns(0).Visible = False
                Else
                    .Filter = ""
                    Set DataGrid2.DataSource = RsFamiliar
                    .MoveFirst
                    DataGrid2.Columns(1).Width = 4095
                    DataGrid2.Columns(2).Width = 1935
                    DataGrid2.Columns(3).Width = 1935
                    DataGrid2.Columns(0).Visible = False
                End If
            End With
    End Select
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    Unload Me
    Form1.Enabled = True
End Sub
Private Sub Text1_Change()
    On Error Resume Next
    With RSTrabajador
        .Requery
        If OPTION1.Value = True Then
            .Filter = "NOMBRE LIKE '*" & Text1 & "*'"
        Else
            .Filter = ""
            Set DataGrid1.DataSource = RSTrabajador
            .MoveFirst
        End If
    End With
End Sub
