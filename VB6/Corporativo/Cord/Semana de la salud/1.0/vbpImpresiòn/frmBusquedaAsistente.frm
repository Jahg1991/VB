VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Persona"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBusquedaAsistente.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   4560
      Picture         =   "frmBusquedaAsistente.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   2760
      Picture         =   "frmBusquedaAsistente.frx":3D4B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1095
      Left            =   2020
      TabIndex        =   2
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
            LCID            =   3082
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
            LCID            =   3082
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
      BackColor       =   &H00C0FFFF&
      Height          =   360
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   0
      Picture         =   "frmBusquedaAsistente.frx":48B1
      Top             =   0
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CnSr As New ADODB.Connection
Private RsSr As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Form1.Enabled = True
            Form1.Text1(0).Text = DataGrid1.Columns(1).Text
            Form1.Text1(1) = DataGrid1.Columns(0).Text
            Set Form1.Text2(0).DataSource = Rs
            Set Form1.Text2(1).DataSource = Rs
            Set Form1.Text2(2).DataSource = Rs
            Set Form1.Text2(3).DataSource = Rs
            Set Form1.Text2(4).DataSource = Rs
            Set Form1.Check1.DataSource = Rs
            Set Form1.Text2(5).DataSource = Rs
            Set Form1.Text2(6).DataSource = Rs
            Set Form1.Text3(0).DataSource = Rs
            Set Form1.Text3(1).DataSource = Rs
            Set Form1.Text3(2).DataSource = Rs
            Set Form1.Text3(3).DataSource = Rs
            Set Form1.Check2(0).DataSource = Rs
            Set Form1.Check2(1).DataSource = Rs
            Set Form1.Text4.DataSource = Rs
            Set Form1.Check3.DataSource = Rs
            Set Form1.Text5.DataSource = Rs
            Set Form1.Check4(0).DataSource = Rs
            Set Form1.Check4(1).DataSource = Rs
            Set Form1.Text6.DataSource = Rs
            Set Form1.Check6.DataSource = Rs
            Set Form1.Text8.DataSource = Rs
            Set Form1.Check5.DataSource = Rs
            Set Form1.Text7.DataSource = Rs
            Set Form1.Check7.DataSource = Rs
            Set Form1.Text9.DataSource = Rs
            Set Form1.Check8(0).DataSource = Rs
            Set Form1.Check8(1).DataSource = Rs
            Set Form1.Text10.DataSource = Rs
            Set Form1.Check9.DataSource = Rs
            Set Form1.Text11.DataSource = Rs
            Set Form1.Check10.DataSource = Rs
            Set Form1.Text12.DataSource = Rs
            Form1.Text2(0).DataField = ("Fecha_nacimiento")
            Form1.Text2(1).DataField = ("Genero")
            Form1.Text2(2).DataField = ("Peso")
            Form1.Text2(3).DataField = ("Talla")
            Form1.Text2(4).DataField = ("Tension_arterial")
            Form1.Check1.DataField = ("Vacuna_toxoide")
            Form1.Text2(5).DataField = ("Otras_vacunas")
            Form1.Text2(6).DataField = ("Observaciones_somatometria")
            Form1.Text3(0).DataField = ("Colesterol")
            Form1.Text3(1).DataField = ("Trigliceridos")
            Form1.Text3(2).DataField = ("Glucosa")
            Form1.Text3(3).DataField = ("Observaciones_laboratorio")
            Form1.Check2(0).DataField = ("Lavado_oidos")
            Form1.Check2(1).DataField = ("Prueba_audicion")
            Form1.Text4.DataField = ("Observaciones_audiometria")
            Form1.Check3.DataField = ("Asistencia_cardiologia")
            Form1.Text5.DataField = ("Cardiologia")
            Form1.Check4(0).DataField = ("Limpieza_dental")
            Form1.Check4(1).DataField = ("Revision_dental")
            Form1.Text6.DataField = ("Observaciones_dental")
            Form1.Check6.DataField = ("Asistencia_Doccu")
            Form1.Text8.DataField = ("Doccu")
            Form1.Check5.DataField = ("Asistencia_docm")
            Form1.Text7.DataField = ("Docm")
            Form1.Check7.DataField = ("Asistencia_mastografia")
            Form1.Text9.DataField = ("mastografia")
            Form1.Check8(0).DataField = ("Consulta_nutricion")
            Form1.Check8(1).DataField = ("Platica_nutricion")
            Form1.Text10.DataField = ("Observaciones_nutricion")
            Form1.Check9.DataField = ("Asistencia_optometria")
            Form1.Text11.DataField = ("Observaciones_optometria")
            Form1.Check10.DataField = ("Asistencia_tuberculosis")
            Form1.Text12.DataField = ("Observaciones_tuberculosis")
            RsSr.Filter = ""
            CnSr.Close
            Unload Form2
        Case 1
            Form1.Enabled = True
            RsSr.Filter = ""
            CnSr.Close
            Unload Form2
    End Select
End Sub

Private Sub DataGrid1_DblClick()
    Text1(0).Text = DataGrid1.Columns(1).Text
    Label2 = DataGrid1.Columns(0).Text
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Text1(0).Text = DataGrid1.Columns(1).Text
        Label2 = DataGrid1.Columns(0).Text
    End If
End Sub

Private Sub Form_Load()
    With CnSr
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=Santateresa1;Persist Security Info=True;User ID=ss16;Initial Catalog=ss16;Data Source=SQLSERVER\SQLEXPRESS;"
    End With
    With RsSr
        If .State = 1 Then .Close
            .Open "select id, nombre from Search", CnSr, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set DataGrid1.DataSource = RsSr
    DataGrid1.Columns(0).DataField = "id"
    DataGrid1.Columns(1).DataField = "nombre"
    DataGrid1.Columns(0).Width = 700
    DataGrid1.Columns(1).Width = 5200
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
        Case 0
            On Error Resume Next
            With RsSr
                .Requery
                If option1.Value = True Then
                    .Filter = "nombre like '*" & Text1(0) & "*'"
                Else
                    .Filter = ""
                    Set DataGrid1.DataSource = RsSr
                    .MoveFirst
                End If
            End With
            DataGrid1.Columns(0).Width = 700
            DataGrid1.Columns(1).Width = 5200
    End Select
End Sub

