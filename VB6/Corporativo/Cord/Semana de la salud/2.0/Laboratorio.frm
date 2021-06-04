VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FormLaboratorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorio"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "Laboratorio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   10335
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
         Height          =   390
         Index           =   5
         Left            =   2160
         TabIndex        =   17
         Top             =   3120
         Width           =   7935
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
         Height          =   390
         Index           =   4
         Left            =   2160
         TabIndex        =   16
         Top             =   2640
         Width           =   855
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
         Height          =   390
         Index           =   3
         Left            =   2160
         TabIndex        =   15
         Top             =   2160
         Width           =   855
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
         Height          =   390
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   1680
         Width           =   855
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
         Height          =   390
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4560
         TabIndex        =   7
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Observaciones"
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
         Index           =   13
         Left            =   360
         TabIndex        =   18
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
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
         Height          =   375
         Index           =   16
         Left            =   2160
         TabIndex        =   12
         Top             =   720
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
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
         Height          =   375
         Index           =   15
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PSA"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Glucosa"
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
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Trigliceridos"
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
         Left            =   720
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Colesterol"
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
         Left            =   960
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Id Asistente"
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
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   1800
      TabIndex        =   10
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
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   14
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Cancelar 
         Caption         =   "Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FormLaboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    On Error Resume Next
    FormLaboratorio.Text1(0).Text = ""
    FormLaboratorio.Text1(1).Text = ""
    FormLaboratorio.Text1(2).Text = ""
    FormLaboratorio.Text1(3).Text = ""
    FormLaboratorio.Text1(4).Text = ""
    FormLaboratorio.Text1(5).Text = ""
    RsNombre.Filter = ""
    RsNombre.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    If Text1(2).Text > 200 Or Text1(0).Text > 160 Or Text1(3).Text > 120 Or Text1(4).Text > 4 Then
        MsgBox "Necesita pasar al módulo de nutrición", vbOKOnly, "Atención"
    End If
    With RsLaboratorio
        .Requery
        .AddNew
            .Fields("ID_AST") = FormLaboratorio.Label1(15).Caption
            .Fields("COLEST") = FormLaboratorio.Text1(2).Text
            .Fields("TRIGLI") = FormLaboratorio.Text1(0).Text
            .Fields("GLUCOS") = FormLaboratorio.Text1(3).Text
            .Fields("PSA") = FormLaboratorio.Text1(4).Text
            .Fields("OBSERV") = FormLaboratorio.Text1(5).Text
        .Update
        .Requery
    End With
    MsgBox ("Información guardada con éxito"), vbOKOnly, "Completado"
    FormLaboratorio.Text1(0).Text = ""
    FormLaboratorio.Text1(1).Text = ""
    FormLaboratorio.Text1(2).Text = ""
    FormLaboratorio.Text1(3).Text = ""
    FormLaboratorio.Text1(4).Text = ""
    FormLaboratorio.Text1(5).Text = ""
    RsNombre.Filter = ""
    RsNombre.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RsLaboratorio
        If .State = 1 Then .Close
            .Open "Select * from LABOR", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsNombre
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RSSexo
        If .State = 1 Then .Close
            .Open "Select ID_AST, GENERO from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set DataGrid1.DataSource = RsNombre
    Set FormLaboratorio.Label1(15).DataSource = RsNombre
    Set FormLaboratorio.Label1(16).DataSource = RsNombre
    FormLaboratorio.Label1(15).DataField = ("ID_AST")
    FormLaboratorio.Label1(16).DataField = ("NOMBRE")
    Set Label2.DataSource = RSSexo
    Label2.DataField = ("GENERO")
    Text1(1).SetFocus
End Sub
Private Sub Guardar_Click()
    On Error Resume Next
    With RsLaboratorio
        .Requery
        .AddNew
            .Fields("ID_AST") = FormLaboratorio.Label1(15).Caption
            .Fields("COLEST") = FormLaboratorio.Text1(2).Text
            .Fields("TRIGLI") = FormLaboratorio.Text1(0).Text
            .Fields("GLUCOS") = FormLaboratorio.Text1(3).Text
            .Fields("PSA") = FormLaboratorio.Text1(4).Text
            .Fields("OBSERV") = FormLaboratorio.Text1(5).Text
        .Update
        .Requery
    End With
    MsgBox ("Información guardada con éxito"), vbOKOnly, "Completado"
    FormLaboratorio.Text1(0).Text = ""
    FormLaboratorio.Text1(1).Text = ""
    FormLaboratorio.Text1(2).Text = ""
    FormLaboratorio.Text1(3).Text = ""
    FormLaboratorio.Text1(4).Text = ""
    FormLaboratorio.Text1(5).Text = ""
    RsNombre.Filter = ""
    RsNombre.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub
Private Sub Label1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 15
            Set Label2.DataSource = RSSexo
            Label2.DataField = ("GENERO")
            With RSSexo
                .Requery
                    If OPTION1.Value = True Then
                        .Filter = "ID_AST LIKE '*" & Label1(15) & "*'"
                    Else
                        .Filter = ""
                        .MoveFirst
                    End If
            End With
    End Select
End Sub
Private Sub Label2_Change()
    On Error Resume Next
    If Label2.Caption = "Femenino" Then
        Text1(4).Enabled = False
        Text1(4).Text = ""
    Else
        Text1(4).Enabled = True
    End If
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    FormLaboratorio.Text1(0).Text = ""
    FormLaboratorio.Text1(1).Text = ""
    FormLaboratorio.Text1(2).Text = ""
    FormLaboratorio.Text1(3).Text = ""
    FormLaboratorio.Text1(4).Text = ""
    FormLaboratorio.Text1(5).Text = ""
    Unload Me
    Form1.Enabled = True
End Sub
Private Sub Text1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1
            With RsNombre
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "NOMBRE LIKE '*" & Text1(1) & "*'"
                Else
                    .Filter = ""
                    Set DataGrid1.DataSource = RsNombre
                    .MoveFirst
                End If
            End With
            Set FormLaboratorio.Label1(15).DataSource = RsNombre
            Set FormLaboratorio.Label1(16).DataSource = RsNombre
            FormLaboratorio.Label1(15).DataField = ("ID_AST")
            FormLaboratorio.Label1(16).DataField = ("NOMBRE")
    End Select
End Sub
