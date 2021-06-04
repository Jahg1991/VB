VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Compras"
   ClientHeight    =   5070
   ClientLeft      =   2505
   ClientTop       =   2550
   ClientWidth     =   10215
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3840
      Top             =   6840
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   630
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   285
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5640
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "No. ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Artìculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   0
      Picture         =   "Form1.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
   Begin VB.Menu Catalogos 
      Caption         =   "Catalogos"
      Begin VB.Menu Proveedores 
         Caption         =   "Proveedores"
         Shortcut        =   ^P
      End
      Begin VB.Menu Articulos 
         Caption         =   "Artìculos"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu Compras2 
      Caption         =   "Compras"
      Begin VB.Menu NuevaCompra 
         Caption         =   "Nueva compra"
         Shortcut        =   ^N
      End
      Begin VB.Menu Historialcompras 
         Caption         =   "Historial de compras"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Articulos_Click()
    OCompra
    CArticulos
End Sub

Private Sub Command1_Click()
    If Text1(0) <> "" And Text2(0) <> "" And Text2(1) <> "" And Text2(2) <> "" And Text2(3) <> "" Then
        ICompra
    End If
End Sub

Private Sub Command2_Click()
    MBProveedor
End Sub

Private Sub Command3_Click()
    MBArticulo
End Sub

Private Sub Command5_Click()
    LeerPuertoBascula
    Command2.Enabled = True
End Sub

Private Sub Compras2_Click()
    'Form3.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
    With RS1
        If .State = 1 Then .Close
            .Open "Select * from articulos", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RS2
        If .State = 1 Then .Close
            .Open "Select * from proveedores", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RS3
        If .State = 1 Then .Close
            .Open "Select * from compras", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RS4
        If .State = 1 Then .Close
            .Open "Select * from r_compras", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
End Sub

Private Sub Historialcompras_Click()
    OCompra
    Form3.Show
End Sub

Private Sub NuevaCompra_Click()
    NCompra
End Sub

Private Sub Proveedores_Click()
    OCompra
    CProveedores
End Sub

Private Sub Salir_Click()
    On Error Resume Next
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    RS1.Close
    RS2.Close
    RS3.Close
    RS4.Close
    CN.Close
    aa = Shell("shutdown -s -t 00")
    Unload Form1
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
        Case 0
            With Form1
                If Text2(1) <> "" And Text1(0) <> "" Then
                        CTotal
                End If
                If Text1(0) = "" Then
                    Command2.Enabled = False
                End If
            End With
    End Select
End Sub

Private Sub Text2_Change(Index As Integer)
    Select Case Index
        Case 0
            If Text2(0) <> "" Then
                Command3.Enabled = True
            End If
        Case 1
            With Form1
                If Text2(1) <> "" Then
                    Text2(2).Enabled = True
                Else
                    Text2(2).Enabled = False
                End If
            End With
        Case 2
            With Form1
                If Text2(1) <> "" And Text1(0) <> "" Then
                    CTotal
                    .Text2(4).Enabled = True
                Else
                    .Text2(4).Enabled = False
                End If
            End With
    End Select
End Sub

Private Sub Timer1_Timer()
    LeerPuertoBascula
End Sub
