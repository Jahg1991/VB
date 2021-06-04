VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PigSale v.1.0 - Ingresar al sistema"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1695
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes:"
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios:"
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas:"
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   1
         Left            =   3000
         Picture         =   "PigSale v.1.0 - Ingresar al sistema.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   0
         Left            =   1320
         Picture         =   "PigSale v.1.0 - Ingresar al sistema.frx":068B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   495
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   1
         Left            =   120
         Picture         =   "PigSale v.1.0 - Ingresar al sistema.frx":0E9A
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "PigSale v.1.0 - Ingresar al sistema.frx":174E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
        
        Case 0
            If Text1(0) = Text1(2).Text And Text1(1) = Text1(3).Text Then
                MsgBox "Bienvenido(a) " & Text1(0), , "PigSale v.1.0"
                Form2.Show
                Form2.Label1.Caption = Text1(0)
                Unload Me
            Else
                MsgBox "El usuario o contraseña no son válidos. Vuelva a intentarlo", , "PigSale v.1.0 - Advertencia"
                Text1(1).SetFocus
                Text1(1).Text = ""
            End If
        
        Case 1
            Unload Me
            Unload Form2
    
    End Select
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    With RsUsers
        If .State = 1 Then .Close
           .Open "Select * from USERS", CnDb, adOpenStatic, adLockOptimistic
           .Requery
    End With
    
    Form2.Command1(0).Enabled = False
    Form2.Command1(1).Enabled = False
    Form2.Command1(2).Enabled = False
        
    Set Text1(2).DataSource = RsUsers
    Set Text1(3).DataSource = RsUsers
    Set Check1(0).DataSource = RsUsers
    Set Check1(1).DataSource = RsUsers
    Set Check1(2).DataSource = RsUsers
    
    Text1(2).DataField = ("NOMBRE")
    Text1(3).DataField = ("PASS")
    Check1(0).DataField = ("VENTAS")
    Check1(1).DataField = ("REPORTES")
    Check1(2).DataField = ("USUARIOS")
    
    If Check1(0).Value = 1 Then
        Form2.Command1(0).Enabled = True
    Else
        Form2.Command1(0).Enabled = False
    End If
    
    If Check1(1).Value = 1 Then
        Form2.Command1(1).Enabled = True
    Else
        Form2.Command1(1).Enabled = False
    End If
    
    If Check1(2).Value = 1 Then
        Form2.Command1(2).Enabled = True
    Else
        Form2.Command1(2).Enabled = False
    End If
    
End Sub

Private Sub Text1_Change(Index As Integer)

    On Error Resume Next

    Select Case Index
    
        Case 0
            With RsUsers
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "NOMBRE LIKE '*" & txtUserName & "*'"
                    If Check1(0).Value = 1 Then
                        Form2.Command1(0).Enabled = True
                    Else
                        Form2.Command1(0).Enabled = False
                    End If
                    If Check1(1).Value = 1 Then
                        Form2.Command1(1).Enabled = True
                    Else
                        Form2.Command1(1).Enabled = False
                    End If
                    If Check1(2).Value = 1 Then
                        Form2.Command1(2).Enabled = True
                    Else
                        Form2.Command1(2).Enabled = False
                    End If
                Else
                    .Filter = ""
                    .MoveFirst
                End If
            End With
        
    End Select

End Sub
