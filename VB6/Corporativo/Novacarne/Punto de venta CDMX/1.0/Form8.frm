VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
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
   LinkTopic       =   "Form8"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   13500
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5535
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   13215
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   7
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   11655
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H009EC0C2&
            Caption         =   "Guardar"
            Height          =   375
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   4800
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5535
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   13215
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "Guardar"
            Height          =   375
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   4800
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   3
            Left            =   1800
            TabIndex        =   11
            Top             =   2160
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   10
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   1800
            MaxLength       =   9
            TabIndex        =   9
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   1
            Left            =   1800
            TabIndex        =   8
            Top             =   720
            Width           =   11055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   0
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UDM"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   6
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Iva"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" And Text1(3) <> "" Then
        With RsItems
            .Requery
            .AddNew
                .Fields("Codigo") = Text1(0)
                .Fields("Descripcion") = Text1(1)
                .Fields("Precio") = Text1(2)
                .Fields("Iva") = Check1.Value
                .Fields("UDM") = Text1(3)
            .Update
            .Requery
        End With
        Text1(0) = ""
        Text1(1) = ""
        Text1(2) = ""
        Text1(3) = ""
        Check1.Value = 0
        Text1(0).SetFocus
    Else
        MsgBox "Favor de llenar todos los campos", vbOKOnly, "Advertencia"
    End If

End Sub

Private Sub Command2_Click()

    If Text1(7) <> "" Then
        With RsClientes
            .Requery
            .AddNew
                .Fields("Nombre") = Text1(7)
            .Update
            .Requery
        End With
        Text1(7) = ""
        Text1(7).SetFocus
    Else
        MsgBox "Favor de llenar todos los campos", vbOKOnly, "Advertencia"
    End If

End Sub

Private Sub Form_Load()

    If TipoCatalogo = 0 Then
        Form8.Caption = "Añadir Artículo"
        Form8.Icon = LoadPicture(App.Path & "\Images\Articulos.ico")
        Frame2.Visible = True
        Frame3.Visible = False
    End If
    If TipoCatalogo = 1 Then
        Form8.Caption = "Añadir Cliente"
        Form8.Icon = LoadPicture(App.Path & "\Images\Clientes.ico")
        Frame2.Visible = False
        Frame3.Visible = True
    End If
    
End Sub
