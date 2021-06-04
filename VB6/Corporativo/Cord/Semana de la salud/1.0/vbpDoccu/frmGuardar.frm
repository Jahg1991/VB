VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guardar"
   ClientHeight    =   1590
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
   ScaleHeight     =   1590
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
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
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
Private Dp As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Form1.Text1(0).Text = "" Then
                MsgBox "No se ha seleccionado a ninguna persona", 64, "Advertencia"
                Form1.Enabled = True
                Cn.Close
                Unload Form4
            Else
                If Label2 = Form1.Label2 Then
                    MsgBox "Esta persona ya està registrada en este mòdulo", 64, "Advertencia"
                    Form1.Enabled = True
                    Cn.Close
                    Unload Form4
                    Exit Sub
                Else
                    With Rs
                        .Requery
                        .AddNew
                            .Fields("Id") = Form1.Label2
                            .Fields("Asistencia") = Form1.Check1(0)
                            .Fields("Observacion") = Form1.Text1(8)
                        .Update
                        .Requery
                    End With
                    Form1.Enabled = True
                    Form1.Text1(0).Text = ""
                    Form1.Text1(8).Text = ""
                    Form1.Check1(0).Value = 0
                    Form1.Label2 = ""
                    Cn.Close
                    Unload Form4
                End If
            End If
        Case 1
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
    With Rs
        If .State = 1 Then .Close
            .Open "select * from doccu", Cn, adOpenDynamic, adLockOptimistic
            .Requery
    End With
    With Dp
        If .State = 1 Then .Close
            If Form1.Label2 <> "" Then
                .Open "select * from doccu where Id = " & Form1.Label2.Caption & "", Cn, adOpenDynamic, adLockOptimistic
            End If
    End With
    Set Label2.DataSource = Dp
    Label2.DataField = "Id"
End Sub


