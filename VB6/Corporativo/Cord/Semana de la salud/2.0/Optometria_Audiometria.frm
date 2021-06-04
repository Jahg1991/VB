VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FormOptometria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optometría"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "Optometria_Audiometria.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
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
      TabIndex        =   5
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      Height          =   2895
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
         TabIndex        =   12
         Top             =   1560
         Width           =   7935
      End
      Begin VB.ComboBox Combo1 
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
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
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
         Left            =   4680
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Asistencia"
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
         Index           =   3
         Left            =   720
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
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
         TabIndex        =   9
         Top             =   1680
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
         TabIndex        =   8
         Top             =   720
         Width           =   8055
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
         TabIndex        =   7
         Top             =   240
         Width           =   1575
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
         Height          =   255
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
      TabIndex        =   6
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
      TabIndex        =   4
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
Attribute VB_Name = "FormOptometria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    On Error Resume Next
    FormOptometria.Text1(1).Text = ""
    FormOptometria.Text1(5).Text = ""
    FormOptometria.Combo1(0).Text = ""
    RsNombre.Filter = ""
    RsNombre.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub

Private Sub Combo1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            If Combo1(0).Text = "Si" Then
                Text1(5).Enabled = True
            Else
                Text1(5).Enabled = False
            End If
    End Select
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    With RsOptometriaAudiometria
        .Requery
        .AddNew
            .Fields("ID_AST") = FormOptometria.Label1(15).Caption
            .Fields("OPTOME") = FormOptometria.Combo1(0).Text
            .Fields("OBSERV") = FormOptometria.Text1(5).Text
        .Update
        .Requery
    End With
    MsgBox ("Información guardada con éxito"), vbOKOnly, "Completado"
    FormOptometria.Text1(1).Text = ""
    FormOptometria.Text1(5).Text = ""
    FormOptometria.Combo1(0).Text = ""
    RsNombre.Filter = ""
    RsNombre.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Combo1(0).AddItem "Si"
    Combo1(0).AddItem "No"
    With RsOptometriaAudiometria
        If .State = 1 Then .Close
            .Open "Select * from OP_AD", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsNombre
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set DataGrid1.DataSource = RsNombre
    Set FormOptometria.Label1(15).DataSource = RsNombre
    Set FormOptometria.Label1(16).DataSource = RsNombre
    FormOptometria.Label1(15).DataField = ("ID_AST")
    FormOptometria.Label1(16).DataField = ("NOMBRE")
    Text1(1).SetFocus
End Sub
Private Sub Guardar_Click()
    On Error Resume Next
    With RsOptometriaAudiometria
        .Requery
        .AddNew
            .Fields("ID_AST") = FormOptometria.Label1(15).Caption
            .Fields("OPTOME") = FormOptometria.Combo1(0).Text
            .Fields("OBSERV") = FormOptometria.Text1(5).Text
        .Update
        .Requery
    End With
    MsgBox ("Información guardada con éxito"), vbOKOnly, "Completado"
    FormOptometria.Text1(1).Text = ""
    FormOptometria.Text1(5).Text = ""
    FormOptometria.Combo1(0).Text = ""
    RsNombre.Filter = ""
    RsNombre.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    FormOptometria.Text1(1).Text = ""
    FormOptometria.Text1(5).Text = ""
    FormOptometria.Combo1(0).Text = ""
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
            Set FormOptometria.Label1(15).DataSource = RsNombre
            Set FormOptometria.Label1(16).DataSource = RsNombre
            FormOptometria.Label1(15).DataField = ("ID_AST")
            FormOptometria.Label1(16).DataField = ("NOMBRE")
    End Select
End Sub
