VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormSomatometriaNuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Somatometría"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "SomatometriaNuevo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   30
         Top             =   3120
         Visible         =   0   'False
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
         Left            =   4440
         TabIndex        =   28
         Top             =   7080
         Width           =   1215
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
         TabIndex        =   27
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
         Index           =   6
         Left            =   3840
         TabIndex        =   26
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
         TabIndex        =   25
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
         Index           =   5
         Left            =   3840
         TabIndex        =   24
         Top             =   5040
         Width           =   1695
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
         TabIndex        =   23
         Top             =   4560
         Width           =   1095
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
         Left            =   3840
         TabIndex        =   22
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1680
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
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
         Format          =   106496001
         CurrentDate     =   42225
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
         Index           =   1
         Left            =   3840
         TabIndex        =   16
         Top             =   720
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
         Index           =   0
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Height          =   495
         Left            =   360
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   5760
         TabIndex        =   31
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
         Index           =   15
         Left            =   3840
         TabIndex        =   29
         Top             =   3600
         Width           =   2775
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
         Index           =   14
         Left            =   3840
         TabIndex        =   21
         Top             =   3120
         Width           =   1695
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
         Height          =   495
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
Attribute VB_Name = "FormSomatometriaNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    On Error Resume Next
    FormSomatometriaNuevo.Text1(0).Text = ""
    FormSomatometriaNuevo.Text1(1).Text = ""
    FormSomatometriaNuevo.DTPicker1.Value = Date
    FormSomatometriaNuevo.Combo1(0).Text = ""
    FormSomatometriaNuevo.Combo1(1).Text = ""
    FormSomatometriaNuevo.Text1(2).Text = ""
    FormSomatometriaNuevo.Label1(14).Caption = ""
    FormSomatometriaNuevo.Label1(15).Caption = ""
    FormSomatometriaNuevo.Text1(3).Text = ""
    FormSomatometriaNuevo.Text1(4).Text = ""
    FormSomatometriaNuevo.Text1(5).Text = ""
    FormSomatometriaNuevo.Combo1(3).Text = ""
    FormSomatometriaNuevo.Text1(6).Text = ""
    FormSomatometriaNuevo.Text1(7).Text = ""
    FormSomatometriaNuevo.Label2.Caption = ""
    Unload BuscarTrabajador
    Command2.Enabled = False
    Command2.Visible = False
    Set Label3.DataSource = RSIDNUM
    Label3.DataField = ("ID_AST")
    RSIDNUM.MoveLast
    Text1(0).SetFocus
End Sub
Private Sub Combo1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1
            If Combo1(1).Text = "Si" Then
                Text1(2).Enabled = True
                Label1(14).Caption = ""
                Label1(15).Caption = ""
                Command2.Enabled = False
                Command2.Visible = False
                BuscarTrabajador.Text1.Text = ""
                BuscarTrabajador.Combo1.Text = ""
                RSTrabajador.Filter = ""
                RSTrabajador.MoveFirst
                Set BuscarTrabajador.DataGrid1.DataSource = RSTrabajador
            Else
                If Combo1(1).Text = "No" Then
                    Text1(2).Text = ""
                    Text1(2).Enabled = False
                    Command2.Enabled = True
                    Command2.Visible = True
                    BuscarTrabajador.Text1.Text = ""
                    BuscarTrabajador.Combo1.Text = ""
                    RSTrabajador.Filter = ""
                    RSTrabajador.MoveFirst
                    Set BuscarTrabajador.DataGrid1.DataSource = RSTrabajador
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
                Label1(14).Caption = ""
                Label1(15).Caption = ""
                Command2.Enabled = False
                Command2.Visible = False
                BuscarTrabajador.Text1.Text = ""
                BuscarTrabajador.Combo1.Text = ""
                RSTrabajador.Filter = ""
                RSTrabajador.MoveFirst
                Set BuscarTrabajador.DataGrid1.DataSource = RSTrabajador
            Else
                If Combo1(1).Text = "No" Then
                    Text1(2).Text = ""
                    Text1(2).Enabled = False
                    Command2.Enabled = True
                    Command2.Visible = True
                    BuscarTrabajador.Text1.Text = ""
                    BuscarTrabajador.Combo1.Text = ""
                    RSTrabajador.Filter = ""
                    RSTrabajador.MoveFirst
                    Set BuscarTrabajador.DataGrid1.DataSource = RSTrabajador
                End If
            End If
    End Select
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    If Not Text1(0) = "" And Not Text1(1) = "" Then
        With RsSomatometria
            .Requery
            .AddNew
                .Fields("ID_AST") = FormSomatometriaNuevo.Text1(0).Text
                .Fields("NOMBRE") = FormSomatometriaNuevo.Text1(1).Text
                .Fields("FE_NAC") = FormSomatometriaNuevo.DTPicker1.Value
                .Fields("GENERO") = FormSomatometriaNuevo.Combo1(0).Text
                .Fields("TRAB_E") = FormSomatometriaNuevo.Combo1(1).Text
                .Fields("AREA_T") = FormSomatometriaNuevo.Text1(2).Text
                .Fields("ID_EMP") = FormSomatometriaNuevo.Label1(14).Caption
                .Fields("PARENT") = FormSomatometriaNuevo.Label1(15).Caption
                .Fields("PES_KG") = FormSomatometriaNuevo.Text1(3).Text
                .Fields("TAL_MT") = FormSomatometriaNuevo.Text1(4).Text
                .Fields("TA") = FormSomatometriaNuevo.Text1(5).Text
                .Fields("VAC_TX") = FormSomatometriaNuevo.Combo1(3).Text
                .Fields("VAC_OT") = FormSomatometriaNuevo.Text1(6).Text
                .Fields("OBSERV") = FormSomatometriaNuevo.Text1(7).Text
                .Fields("EDAD") = FormSomatometriaNuevo.Label2.Caption
            .Update
        End With
        With RSIDNUM
            .Requery
            .AddNew
                .Fields("ID_AST") = Text1(0).Text
            .Update
            .Requery
        End With
        MsgBox ("Asistente " & FormSomatometriaNuevo.Text1(1) & " registrado"), vbOKOnly, "Completado"
        FormSomatometriaNuevo.Text1(0).Text = ""
        FormSomatometriaNuevo.Text1(1).Text = ""
        FormSomatometriaNuevo.DTPicker1.Value = Date
        FormSomatometriaNuevo.Combo1(0).Text = ""
        FormSomatometriaNuevo.Combo1(1).Text = ""
        FormSomatometriaNuevo.Text1(2).Text = ""
        FormSomatometriaNuevo.Label1(14).Caption = ""
        FormSomatometriaNuevo.Label1(15).Caption = ""
        FormSomatometriaNuevo.Text1(3).Text = ""
        FormSomatometriaNuevo.Text1(4).Text = ""
        FormSomatometriaNuevo.Text1(5).Text = ""
        FormSomatometriaNuevo.Combo1(3).Text = ""
        FormSomatometriaNuevo.Text1(6).Text = ""
        FormSomatometriaNuevo.Text1(7).Text = ""
        FormSomatometriaNuevo.Label2.Caption = ""
        Command2.Enabled = False
        Command2.Visible = False
        BuscarTrabajador.Text1.Text = ""
        BuscarTrabajador.Combo1.Text = ""
        RSTrabajador.Filter = ""
        RSTrabajador.MoveFirst
        Set BuscarTrabajador.DataGrid1.DataSource = RSTrabajador
        Text1(0).SetFocus
        Set Label3.DataSource = RSIDNUM
        Label3.DataField = ("ID_AST")
        RSIDNUM.MoveLast
    Else
        MsgBox ("Código y nombre necesarios para registrar al asistente"), vbOKOnly, "Error"
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    BuscarTrabajador.Show
    FormSomatometriaNuevo.Enabled = False
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
    With RSTrabajador
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT Where TRAB_E = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RSIDNUM
        If .State = 1 Then .Close
            .Open "Select * from ID", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set Label3.DataSource = RSIDNUM
    Label3.DataField = ("ID_AST")
    RSIDNUM.MoveLast
    Text1(0).Text = Val(Label3.Caption) + 1
    Combo1(0).AddItem "Femenino"
    Combo1(0).AddItem "Masculino"
    Combo1(1).AddItem "Si"
    Combo1(1).AddItem "No"
    Combo1(3).AddItem "Si"
    Combo1(3).AddItem "No"
    DTPicker1.Value = Date
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
    If Not Text1(0) = "" And Not Text1(1) = "" Then
        With RsSomatometria
            .Requery
            .AddNew
                .Fields("ID_AST") = FormSomatometriaNuevo.Text1(0).Text
                .Fields("NOMBRE") = FormSomatometriaNuevo.Text1(1).Text
                .Fields("FE_NAC") = FormSomatometriaNuevo.DTPicker1.Value
                .Fields("GENERO") = FormSomatometriaNuevo.Combo1(0).Text
                .Fields("TRAB_E") = FormSomatometriaNuevo.Combo1(1).Text
                .Fields("AREA_T") = FormSomatometriaNuevo.Text1(2).Text
                .Fields("ID_EMP") = FormSomatometriaNuevo.Label1(14).Caption
                .Fields("PARENT") = FormSomatometriaNuevo.Label1(15).Caption
                .Fields("PES_KG") = FormSomatometriaNuevo.Text1(3).Text
                .Fields("TAL_MT") = FormSomatometriaNuevo.Text1(4).Text
                .Fields("TA") = FormSomatometriaNuevo.Text1(5).Text
                .Fields("VAC_TX") = FormSomatometriaNuevo.Combo1(3).Text
                .Fields("VAC_OT") = FormSomatometriaNuevo.Text1(6).Text
                .Fields("OBSERV") = FormSomatometriaNuevo.Text1(7).Text
                .Fields("EDAD") = FormSomatometriaNuevo.Label2.Caption
            .Update
        End With
        With RSIDNUM
            .Requery
            .AddNew
                .Fields("ID_AST") = Text1(0).Text
            .Update
            .Requery
        End With
        MsgBox ("Asistente " & FormSomatometriaNuevo.Text1(1) & " registrado"), vbOKOnly, "Completado"
        FormSomatometriaNuevo.Text1(0).Text = ""
        FormSomatometriaNuevo.Text1(1).Text = ""
        FormSomatometriaNuevo.DTPicker1.Value = Date
        FormSomatometriaNuevo.Combo1(0).Text = ""
        FormSomatometriaNuevo.Combo1(1).Text = ""
        FormSomatometriaNuevo.Text1(2).Text = ""
        FormSomatometriaNuevo.Label1(14).Caption = ""
        FormSomatometriaNuevo.Label1(15).Caption = ""
        FormSomatometriaNuevo.Text1(3).Text = ""
        FormSomatometriaNuevo.Text1(4).Text = ""
        FormSomatometriaNuevo.Text1(5).Text = ""
        FormSomatometriaNuevo.Combo1(3).Text = ""
        FormSomatometriaNuevo.Text1(6).Text = ""
        FormSomatometriaNuevo.Text1(7).Text = ""
        FormSomatometriaNuevo.Label2.Caption = ""
        Command2.Enabled = False
        Command2.Visible = False
        BuscarTrabajador.Text1.Text = ""
        BuscarTrabajador.Combo1.Text = ""
        RSTrabajador.Filter = ""
        RSTrabajador.MoveFirst
        Set BuscarTrabajador.DataGrid1.DataSource = RSTrabajador
        Text1(0).SetFocus
        Set Label3.DataSource = RSIDNUM
        Label3.DataField = ("ID_AST")
        RSIDNUM.MoveLast
    Else
        MsgBox ("Código y nombre necesarios para registrar al asistente"), vbOKOnly, "Error"
    End If
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    FormSomatometriaNuevo.Text1(0).Text = ""
    FormSomatometriaNuevo.Text1(1).Text = ""
    FormSomatometriaNuevo.DTPicker1.Value = Date
    FormSomatometriaNuevo.Combo1(0).Text = ""
    FormSomatometriaNuevo.Combo1(1).Text = ""
    FormSomatometriaNuevo.Text1(2).Text = ""
    FormSomatometriaNuevo.Label1(14).Caption = ""
    FormSomatometriaNuevo.Label1(15).Caption = ""
    FormSomatometriaNuevo.Text1(3).Text = ""
    FormSomatometriaNuevo.Text1(4).Text = ""
    FormSomatometriaNuevo.Text1(5).Text = ""
    FormSomatometriaNuevo.Combo1(3).Text = ""
    FormSomatometriaNuevo.Text1(6).Text = ""
    FormSomatometriaNuevo.Text1(7).Text = ""
    FormSomatometriaNuevo.Label2.Caption = ""
    BuscarTrabajador.Text1.Text = ""
    BuscarTrabajador.Combo1.Text = ""
    Form1.Enabled = True
    Unload Me
End Sub
