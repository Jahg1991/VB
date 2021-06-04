VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FormSomatometriaRegistrado 
   Caption         =   "Somatometría"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "SomatometriaRegistrado.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9675
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   1800
      TabIndex        =   17
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2355
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
      TabIndex        =   16
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   10335
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Left            =   5640
         TabIndex        =   31
         Top             =   3120
         Width           =   495
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
         TabIndex        =   30
         Top             =   6960
         Width           =   1335
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
         Index           =   8
         Left            =   3840
         TabIndex        =   29
         Top             =   6480
         Width           =   6255
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
         Index           =   7
         Left            =   3840
         TabIndex        =   28
         Top             =   6000
         Width           =   6255
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
         Index           =   3
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   5520
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
         Index           =   6
         Left            =   3840
         TabIndex        =   26
         Top             =   5040
         Width           =   1935
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
         Index           =   5
         Left            =   3840
         TabIndex        =   25
         Top             =   4560
         Width           =   975
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
         Left            =   3840
         TabIndex        =   24
         Top             =   4080
         Width           =   975
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
         Left            =   3840
         TabIndex        =   23
         Top             =   2640
         Width           =   4095
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
         Index           =   1
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2160
         Width           =   855
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
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   224788481
         CurrentDate     =   42225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
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
         Index           =   18
         Left            =   3840
         TabIndex        =   33
         Top             =   3600
         Width           =   1815
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
         Index           =   17
         Left            =   3840
         TabIndex        =   32
         Top             =   3120
         Width           =   1815
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
         Left            =   3840
         TabIndex        =   19
         Top             =   720
         Width           =   6255
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
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   1575
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
         Left            =   2040
         TabIndex        =   14
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Otras vacunas"
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
         Index           =   12
         Left            =   2160
         TabIndex        =   13
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vacuna toxoide"
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
         Index           =   11
         Left            =   2040
         TabIndex        =   12
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Presión arterial"
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
         Index           =   10
         Left            =   2040
         TabIndex        =   11
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Talla en cm"
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
         Index           =   9
         Left            =   2520
         TabIndex        =   10
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso en Kg"
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
         Index           =   8
         Left            =   2400
         TabIndex        =   9
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Parentezco"
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
         Left            =   2520
         TabIndex        =   8
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Familiar que labora en la empresa"
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
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Área de trabajo"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Trabajador de la empresa"
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
         Left            =   960
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Género"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
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
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   2055
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
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
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
      TabIndex        =   15
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
Attribute VB_Name = "FormSomatometriaRegistrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    On Error Resume Next
    FormSomatometriaRegistrado.Text1(1).Text = ""
    RsSomatometria.Filter = ""
    RsSomatometria.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    BuscarTrabajadorReg.Text1.Text = ""
    BuscarTrabajadorReg.Combo1.Text = ""
    RSTrabajador.Filter = ""
    RSTrabajador.MoveFirst
    Set BuscarTrabajadorReg.DataGrid1.DataSource = RSTrabajador
    Text1(1).SetFocus
End Sub
Private Sub Combo1_Change(Index As Integer)
    On Error Resume Next
    On Error Resume Next
    Select Case Index
        Case 1
            If Combo1(1).Text = "Si" Then
                Text1(2).Enabled = True
                Label1(17).Caption = ""
                Label1(18).Caption = ""
                Command2.Enabled = False
                Command2.Visible = False
            Else
                If Combo1(1).Text = "No" Then
                    Text1(2).Enabled = False
                    Text1(1).Text = ""
                    Command2.Enabled = True
                    Command2.Visible = True
                End If
            End If
    End Select
End Sub
Private Sub Combo1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1
            If Combo1(1).Text = "Si" Then
                Text1(2).Enabled = True
                Label1(17).Caption = ""
                Label1(18).Caption = ""
                Command2.Enabled = False
                Command2.Visible = False
            Else
                If Combo1(1).Text = "No" Then
                    Text1(2).Enabled = False
                    Text1(1).Text = ""
                    Command2.Enabled = True
                    Command2.Visible = True
                End If
            End If
    End Select
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    With RsSomatometria
        .Update
    End With
    MsgBox "Información guardada con éxito", vbOKOnly, "Completado"
    FormSomatometriaRegistrado.Text1(1).Text = ""
    RsSomatometria.Filter = ""
    RsSomatometria.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    BuscarTrabajadorReg.Text1.Text = ""
    BuscarTrabajadorReg.Combo1.Text = ""
    RSTrabajador.Filter = ""
    RSTrabajador.MoveFirst
    Set BuscarTrabajadorReg.DataGrid1.DataSource = RSTrabajador
    Text1(1).SetFocus
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    BuscarTrabajadorReg.Show
    Me.Enabled = False
End Sub

Private Sub DTPicker1_Change()
    On Error Resume Next
    Dim Fecha_Nacimiento As Date
    Dim Años As Variant
    Fecha_Nacimiento = DTPicker1.Value
    If IsNull(Fecha_Nacimiento) Then
        Calcular_Edad = 0
    End If
    Años = DateDiff("yyyy", Fecha_Nacimiento, Now)
    If Date < DateSerial(Year(Now), Month(Fecha_Nacimiento), Day(Fecha_Nacimiento)) Then
        Años = Años - 1
    End If
    Calcular_Edad = CInt(Años)
    Label2.Caption = Años
End Sub
Private Sub DTPicker1_Click()
    On Error Resume Next
    Dim Fecha_Nacimiento As Date
    Dim Años As Variant
    Fecha_Nacimiento = DTPicker1.Value
    If IsNull(Fecha_Nacimiento) Then
        Calcular_Edad = 0
    End If
    Años = DateDiff("yyyy", Fecha_Nacimiento, Now)
    If Date < DateSerial(Year(Now), Month(Fecha_Nacimiento), Day(Fecha_Nacimiento)) Then
        Años = Años - 1
    End If
    Calcular_Edad = CInt(Años)
    Label2.Caption = Años
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RsSomatometria
        If .State = 1 Then .Close
            .Open "Select * from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsNombre
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set FormSomatometriaRegistrado.Label1(15).DataSource = RsNombre
    Set FormSomatometriaRegistrado.Label1(16).DataSource = RsNombre
    Set FormSomatometriaRegistrado.DTPicker1.DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Combo1(0).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Combo1(1).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Text1(2).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Label1(17).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Label1(18).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Text1(4).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Text1(5).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Text1(6).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Combo1(3).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Text1(7).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Text1(8).DataSource = RsSomatometria
    Set FormSomatometriaRegistrado.Label2.DataSource = RsSomatometria
    FormSomatometriaRegistrado.Label1(15).DataField = ("ID_AST")
    FormSomatometriaRegistrado.Label1(16).DataField = ("NOMBRE")
    FormSomatometriaRegistrado.DTPicker1.DataField = ("FE_NAC")
    FormSomatometriaRegistrado.Combo1(0).DataField = ("GENERO")
    FormSomatometriaRegistrado.Combo1(1).DataField = ("TRAB_E")
    FormSomatometriaRegistrado.Text1(2).DataField = ("AREA_T")
    FormSomatometriaRegistrado.Label1(17).DataField = ("ID_EMP")
    FormSomatometriaRegistrado.Label1(18).DataField = ("PARENT")
    FormSomatometriaRegistrado.Text1(4).DataField = ("PES_KG")
    FormSomatometriaRegistrado.Text1(5).DataField = ("TAL_MT")
    FormSomatometriaRegistrado.Text1(6).DataField = ("TA")
    FormSomatometriaRegistrado.Combo1(3).DataField = ("VAC_TX")
    FormSomatometriaRegistrado.Text1(7).DataField = ("VAC_OT")
    FormSomatometriaRegistrado.Text1(8).DataField = ("OBSERV")
    FormSomatometriaRegistrado.Label2.DataField = ("EDAD")
    Combo1(0).AddItem "Femenino"
    Combo1(0).AddItem "Masculino"
    Combo1(1).AddItem "Si"
    Combo1(1).AddItem "No"
    Combo1(3).AddItem "Si"
    Combo1(3).AddItem "No"
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
    Dim Fecha_Nacimiento As Date
    Dim Años As Variant
    Fecha_Nacimiento = DTPicker1.Value
    If IsNull(Fecha_Nacimiento) Then
        Calcular_Edad = 0
    End If
    Años = DateDiff("yyyy", Fecha_Nacimiento, Now)
    If Date < DateSerial(Year(Now), Month(Fecha_Nacimiento), Day(Fecha_Nacimiento)) Then
        Años = Años - 1
    End If
    Calcular_Edad = CInt(Años)
    Label2.Caption = Años
End Sub
Private Sub Guardar_Click()
    On Error Resume Next
    With RsSomatometria
        .Update
    End With
    MsgBox "Información guardada con éxito", vbOKOnly, "Completado"
    FormSomatometriaRegistrado.Text1(1).Text = ""
    RsSomatometria.Filter = ""
    RsSomatometria.MoveFirst
    Set DataGrid1.DataSource = RsNombre
    BuscarTrabajadorReg.Text1.Text = ""
    BuscarTrabajadorReg.Combo1.Text = ""
    RSTrabajador.Filter = ""
    RSTrabajador.MoveFirst
    Set BuscarTrabajadorReg.DataGrid1.DataSource = RSTrabajador
    Text1(1).SetFocus
End Sub
Private Sub Label1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 15
            With RsSomatometria
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
Private Sub Salir_Click()
    On Error Resume Next
    FormSomatometriaRegistrado.Text1(1).Text = ""
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
            Set FormSomatometriaRegistrado.Label1(15).DataSource = RsNombre
            Set FormSomatometriaRegistrado.Label1(16).DataSource = RsNombre
            Set FormSomatometriaRegistrado.DTPicker1.DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Combo1(0).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Combo1(1).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Text1(2).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Label1(17).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Label1(18).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Text1(4).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Text1(5).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Text1(6).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Combo1(3).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Text1(7).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Text1(8).DataSource = RsSomatometria
            Set FormSomatometriaRegistrado.Label2.DataSource = RsSomatometria
            FormSomatometriaRegistrado.Label1(15).DataField = ("ID_AST")
            FormSomatometriaRegistrado.Label1(16).DataField = ("NOMBRE")
            FormSomatometriaRegistrado.DTPicker1.DataField = ("FE_NAC")
            FormSomatometriaRegistrado.Combo1(0).DataField = ("GENERO")
            FormSomatometriaRegistrado.Combo1(1).DataField = ("TRAB_E")
            FormSomatometriaRegistrado.Text1(2).DataField = ("AREA_T")
            FormSomatometriaRegistrado.Label1(17).DataField = ("ID_EMP")
            FormSomatometriaRegistrado.Label1(18).DataField = ("PARENT")
            FormSomatometriaRegistrado.Text1(4).DataField = ("PES_KG")
            FormSomatometriaRegistrado.Text1(5).DataField = ("TAL_MT")
            FormSomatometriaRegistrado.Text1(6).DataField = ("TA")
            FormSomatometriaRegistrado.Combo1(3).DataField = ("VAC_TX")
            FormSomatometriaRegistrado.Text1(7).DataField = ("VAC_OT")
            FormSomatometriaRegistrado.Text1(8).DataField = ("OBSERV")
            FormSomatometriaRegistrado.Label2.DataField = ("EDAD")
    End Select
End Sub
