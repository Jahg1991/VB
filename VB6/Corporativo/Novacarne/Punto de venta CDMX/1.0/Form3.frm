VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13575
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1440
            TabIndex        =   13
            Top             =   1680
            Width           =   11895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1440
            MaxLength       =   9
            TabIndex        =   12
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   11
            Top             =   960
            Width           =   11295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1440
            TabIndex        =   10
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H009EC0C2&
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   5040
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1440
            TabIndex        =   8
            Top             =   120
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   10404034
            CalendarTrailingForeColor=   10404034
            Format          =   168755201
            CurrentDate     =   43810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F6E4D3&
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   4
            Left            =   -360
            TabIndex        =   6
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F6E4D3&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   3
            Left            =   -360
            TabIndex        =   5
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F6E4D3&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   -360
            TabIndex        =   4
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F6E4D3&
            BackStyle       =   0  'Transparent
            Caption         =   "Folio"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   1
            Left            =   -360
            TabIndex        =   3
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F6E4D3&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   0
            Left            =   -360
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Form2.Show
    Form2.Text1.SetFocus
    
End Sub

Private Sub Command2_Click()

    If IdCliente = "" Then
        MsgBox "Elija un cliente", vbOKOnly, "Advertencia"
    Else
        If Text1(2) = "" Then
            MsgBox "Introduzca el total", vbOKOnly, "Advertencia"
        Else
            With RsPagos
                .AddNew
                    .Fields("Fecha") = DTPicker1.Value
                    .Fields("Folio") = Text1(0)
                    .Fields("Cliente") = IdCliente
                    .Fields("Total") = Text1(2)
                    .Fields("Referencia") = Text1(3)
                .Update
                .Requery
            End With
            DTPicker1.Value = Date
            With RsIdPagos
                .Requery
                .MoveFirst
                IdPagos = RsIdPagos!IdPagos
            End With
            Text1(0).Text = "CDMX-" & IdPagos
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            IdCliente = ""
        End If
    End If

End Sub

Private Sub Form_Load()

    DTPicker1.Value = Date
    With RsIdPagos
        .Requery
        .MoveFirst
        IdPagos = RsIdPagos!IdPagos
    End With
    Text1(0).Text = "CDMX-" & IdPagos
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    IdCliente = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    TipoTransaccion = ""
    IdCliente = ""
    Form1.Enabled = True
    
End Sub
