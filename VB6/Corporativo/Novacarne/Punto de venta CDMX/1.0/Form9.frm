VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferencias"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   13695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5535
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13455
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   26
            Top             =   4800
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   2160
            TabIndex        =   25
            Top             =   4440
            Width           =   5055
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   2160
            TabIndex        =   24
            Top             =   4080
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "Guardar"
            Height          =   375
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   5040
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   23
            Top             =   3720
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   22
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   21
            Top             =   3000
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   20
            Top             =   2640
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   2160
            TabIndex        =   19
            Top             =   2280
            Width           =   5055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   2160
            TabIndex        =   18
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   2160
            TabIndex        =   17
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   16
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   2160
            TabIndex        =   15
            Top             =   840
            Width           =   10935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   14
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   13
            Top             =   120
            Width           =   10935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ERP"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   30
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impresora"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   29
            Top             =   4440
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ver Preferencias"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   28
            Top             =   4080
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Imprimir Ticket"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   12
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Modificar precio"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   11
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Añadir Articulos"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   10
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Añadir Clientes"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   9
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Correo electrónico"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   8
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   7
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código Postal"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ciudad"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RFC"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    With RsPreferencias
        .Update
        .Requery
    End With
    
    Unload Form1
    Unload Me

End Sub

Private Sub Form_Load()

    On Error Resume Next
    Set Text1(0).DataSource = RsPreferencias
    Set Text1(1).DataSource = RsPreferencias
    Set Text1(2).DataSource = RsPreferencias
    Set Text1(3).DataSource = RsPreferencias
    Set Text1(4).DataSource = RsPreferencias
    Set Text1(5).DataSource = RsPreferencias
    Set Text1(6).DataSource = RsPreferencias
    Set Text1(7).DataSource = RsPreferencias
    Set Check1(0).DataSource = RsPreferencias
    Set Check1(1).DataSource = RsPreferencias
    Set Check1(2).DataSource = RsPreferencias
    Set Check1(3).DataSource = RsPreferencias
    Set Check1(4).DataSource = RsPreferencias
    Set Check1(5).DataSource = RsPreferencias
    With RsPreferencias
        .Requery
        .MoveFirst
        Text1(0).DataField = .Fields("Empresa")
        Text1(1).DataField = .Fields("RFC")
        Text1(2).DataField = .Fields("Ciudad")
        Text1(3).DataField = .Fields("Estado")
        Text1(4).DataField = .Fields("Codigo Postal")
        Text1(5).DataField = .Fields("Telefono")
        Text1(6).DataField = .Fields("Correo")
        Check1(0).DataField = .Fields("Añadir Clientes")
        Check1(1).DataField = .Fields("Añadir Articulos")
        Check1(2).DataField = .Fields("Modificar Precio Venta")
        Check1(3).DataField = .Fields("Imprimir Ticket")
        Check1(4).DataField = .Fields("Ver Preferencias")
        Text1(7).DataField = .Fields("Impresora de Tickets")
        Check1(5).DataField = .Fields("ERP")
        .Requery
        .MoveFirst
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Form1.Enabled = True

End Sub
