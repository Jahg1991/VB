VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guardar"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
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
   Icon            =   "frmGuardar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   1920
      Picture         =   "frmGuardar.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmGuardar.frx":3D4B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿Desea guardar los cambios?"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Cn As New ADODB.Connection
Private Rs As New ADODB.Recordset
Private RsId As New ADODB.Recordset
Private RsSr As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Form1.Text1(0).Text = "" Then
                MsgBox "No se ha seleccionado a ninguna persona", 64, "Advertencia"
                Form1.Enabled = True
                Form1.Text1(0).SetFocus
                Rs.Close
                RsId.Close
                RsSr.Close
                Cn.Close
                Unload Form4
            Else
                With RsId
                    .Requery
                    .AddNew
                        .Fields("id") = Label2(1)
                    .Update
                    .Requery
                End With
                With Rs
                    .Requery
                    .AddNew
                        .Fields("Id") = Label2(1)
                        .Fields("Nombre") = Form1.Text1(0)
                        .Fields("FechaNacimiento") = Form1.DTPicker1
                        .Fields("Edad") = Form1.Text1(1)
                        .Fields("Genero") = Form1.Combo1(0)
                        .Fields("TrabajadorEmpresa") = Form1.Check1(0)
                        .Fields("Area") = Form1.Text1(2)
                        .Fields("Familiar") = Form1.Text1(3)
                        .Fields("Peso") = Form1.Text1(4)
                        .Fields("Talla") = Form1.Text1(5)
                        .Fields("TensionArterial") = Form1.Text1(6)
                        .Fields("VacunaToxoide") = Form1.Check2(0)
                        .Fields("OtrasVacunas") = Form1.Check2(1)
                        .Fields("Observacion") = Form1.Text1(8)
                        .Fields("NombreVacuna") = Form1.Text1(7)
                    .Update
                    .Requery
                End With
                With RsSr
                    .Requery
                    .AddNew
                        .Fields("id") = Label2(1)
                        .Fields("nombre") = Form1.Text1(0)
                        .Fields("edad") = Form1.Text1(1)
                        .Fields("genero") = Form1.Combo1(0)
                    .Update
                    .Requery
                End With
                Form1.Enabled = True
                Form1.Text1(0).Text = ""
                Form1.Text1(1).Text = ""
                Form1.Text1(2).Text = ""
                Form1.Text1(3).Text = ""
                Form1.Text1(4).Text = ""
                Form1.Text1(5).Text = ""
                Form1.Text1(6).Text = ""
                Form1.Text1(7).Text = ""
                Form1.Text1(8).Text = ""
                Form1.Check1(0).Value = 0
                Form1.Check2(0).Value = 0
                Form1.Check2(1).Value = 0
                Form1.DTPicker1.Value = Date
                Form1.Combo1(0) = "Femenino"
                Form1.Text1(2).Enabled = False
                Form1.Text1(7).Enabled = False
                Rs.Close
                RsId.Close
                RsSr.Close
                Cn.Close
                Unload Form4
            End If
        Case 1
            Rs.Close
            RsId.Close
            RsSr.Close
            Cn.Close
            Form1.Enabled = True
            Unload Form4
    End Select
End Sub

Private Sub Form_Load()
    With Cn
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=Santateresa1;Persist Security Info=True;User ID=ss16;Initial Catalog=ss16;Data Source=SQLSERVER\SQLEXPRESS;"
    End With
    With RsId
        If .State = 1 Then .Close
            .Open "select * from id", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With Rs
        If .State = 1 Then .Close
            .Open "select * from somatometria", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsSr
        If .State = 1 Then .Close
            .Open "select * from Search", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set Label2(0).DataSource = RsId
    Label2(0).DataField = ("id")
    If RsId.EOF <> True Then
        RsId.MoveLast
    End If
    Label2(1) = Label2(0) + 1
End Sub
