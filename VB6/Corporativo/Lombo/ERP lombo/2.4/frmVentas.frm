VERSION 5.00
Begin VB.Form frmVentas 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Nueva Venta"
   ClientHeight    =   9075
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   17415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00404000&
      Height          =   2295
      Index           =   2
      Left            =   6277
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   4860
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   4575
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ACEPTAR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   8
            Left            =   1800
            TabIndex        =   20
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   9
            Left            =   1800
            TabIndex        =   21
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   19
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   6480
      Index           =   0
      Left            =   5100
      TabIndex        =   44
      Top             =   1537
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6255
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   6975
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   2
            Left            =   2040
            TabIndex        =   18
            Top             =   4320
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   1
            Left            =   2040
            TabIndex        =   17
            Top             =   3600
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   0
            Left            =   2040
            TabIndex        =   16
            Top             =   2880
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   3
            Left            =   2040
            TabIndex        =   15
            Top             =   2160
            Width           =   4575
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   540
            Index           =   4
            Left            =   2040
            TabIndex        =   14
            Top             =   1440
            Width           =   4575
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   840
            Width           =   4575
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ACEPTAR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   5640
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CAMBIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   17
            Left            =   240
            TabIndex        =   53
            Top             =   5040
            Width           =   6495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CAMBIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   6
            Left            =   -240
            TabIndex        =   52
            Top             =   4320
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PAGADO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   5
            Left            =   -240
            TabIndex        =   51
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   4
            Left            =   -240
            TabIndex        =   50
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PUNTOS"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   49
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   2
            Left            =   -240
            TabIndex        =   48
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "REFERENCIA"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TERMINAL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   0
            Left            =   -240
            TabIndex        =   46
            Top             =   840
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   17175
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   2
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3600
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   12735
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   7
         Left            =   2400
         TabIndex        =   42
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   0
         Left            =   2400
         TabIndex        =   30
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   1600
         Width           =   14535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   1
         Left            =   2400
         MaxLength       =   7
         TabIndex        =   5
         Top             =   2140
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   2
         Left            =   2400
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2640
         Width           =   3015
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1485
         Left            =   240
         TabIndex        =   11
         Top             =   5520
         Width           =   16695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   3
         Left            =   2400
         TabIndex        =   7
         Top             =   3120
         Width           =   14535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   5
         Left            =   10320
         TabIndex        =   29
         Top             =   7680
         Width           =   6615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   6
         Left            =   10320
         TabIndex        =   28
         Top             =   7200
         Width           =   6615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   4
         Left            =   10320
         TabIndex        =   27
         Top             =   8160
         Width           =   6615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "ELIMINAR"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "A?ADIR"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   33
         Left            =   1200
         TabIndex        =   43
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FOLIO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   32
         Left            =   1200
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   31
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ARTICULO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   30
         Left            =   120
         TabIndex        =   39
         Top             =   1600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   29
         Left            =   120
         TabIndex        =   38
         Top             =   2140
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   28
         Left            =   -360
         TabIndex        =   37
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE VENTA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   27
         Left            =   120
         TabIndex        =   36
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   26
         Left            =   8040
         TabIndex        =   35
         Top             =   8160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COMENTARIOS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   25
         Left            =   120
         TabIndex        =   34
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   24
         Left            =   8040
         TabIndex        =   33
         Top             =   7680
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   23
         Left            =   8040
         TabIndex        =   32
         Top             =   7200
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmVentas.frx":0000
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   22
         Left            =   360
         TabIndex        =   31
         Top             =   5160
         Width           =   12615
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Anadir 
         Caption         =   "A?adir Cliente"
         Shortcut        =   ^A
      End
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmVentas
'Proposito:     Registro de ventas
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Se agrego usuario, fecha de creacion
'                                               y modificacion a todos los insert
'
'1.2        14/05/2021     Alfredo Hernandez    Se agrego validacion para inv.
'                                               negativos
'
'1.3        18/05/2021     Alfredo Hernandez    Se modifico la actualizacion en el
'                                               uso de pedidos
'
'1.4        20/05/2021     Alfredo Hernandez    Se agrego validacion para pago con
'                                               puntos si no se cumple la cantidad
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset    'folio
Dim RS1 As New adodb.Recordset    'clientesproveedores
Dim Rs2 As New adodb.Recordset    'items
Dim Rs3 As New adodb.Recordset    'ventascompras
Dim Rs4 As New adodb.Recordset    'lista de ingredientes
Dim Rs5 As New adodb.Recordset    'movimientos de inventarios
Dim Rs6 As New adodb.Recordset    'ticket
Dim Rs9 As New adodb.Recordset    'movimientos de caja
Dim Rs10 As New adodb.Recordset    'puntos
Dim Rs11 As New adodb.Recordset    'terjeta
Dim Rs12 As New adodb.Recordset    'lotes
Dim Rs13 As New adodb.Recordset    'cabecera
Dim Rs14 As New adodb.Recordset    'tipo de terminal
'//OTROS
Dim TipoErr As Long
Dim i As Long
Dim X As Long
Dim intX As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim c4 As Long
Dim Prt As Printer
'//VALORES PARA INSERTAR
Dim v1 As Long    'idclienteproveedor
Dim v2 As String    'nombre cliente proveedor
Dim v3 As String    'folio
Dim v4 As String    'LugarVenta
Dim v5 As String    'mesa
Dim v6 As Date    'fecha
Dim v7 As String    'Tipoarticulo
Dim v8 As Long    'idarticulo
Dim v9 As String    'codigo articulo
Dim v10 As String    'descripcion articulo
Dim v11 As String    'cantidad
Dim v12 As String    'UDM
Dim v13 As String    'precio
Dim v14 As String    'subtotal
Dim v15 As String    'iva
Dim v16 As String    'total
Dim v17 As String    'totalpagado
Dim v18 As String    'cancelado
Dim v19 As String    'comentarios
Dim v20 As String    'tipo
Dim IdTransaccion As Long           'folio
'//ARTICULOS
Dim InItemId As Long
Dim vCategoria As String
'//CLIENTES
Dim Credito As String   'credito del cliente
Dim CreditoUsado As String        'Credito usado del cliente
Dim DiasCredito As Long         'Dias de credito del cliente
Dim DiasCreditoUsado As Long              'Dias de credito usados por el cliente
Dim ClienteMayorista As String            '?Es cliente mayorista?
Dim ListaPrecios As Long          'lista de precios cliente
'//LOTE
Dim ControlLote As Boolean
Dim CantidadRestante As String
Dim vLote As String
Dim vCantidadLote As String
Dim vCurrentLote As String
Dim InLoteExiste As Long
'//VENTAS
Dim listSubtotal As String
Dim listIva As String
Dim listTotal As String
Dim vLstCantidad As String
Dim vLstPrecio As String
Dim viva As String
Dim viid As String
Dim videscripcion As String
Dim vicantidad As String
Dim viprecio As String
Dim Array_Comentarios() As String
'//PAGOS
Dim DineroRestante As String
'//TICKET
Dim vTicketSubtotal As String
Dim vTicketIva As String
Dim vTicketTotal As String
'//PEDIDOS
Dim sql As String

Private Sub Form_Load()
    On Error GoTo errHandler
    If StPermisosCatalogos = "Si" Then
        With Anadir
            .Visible = True
        End With
    Else
        With Anadir
            .Visible = False
        End With
    End If

    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With Rs
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from PO_TRANSACTION_ID_R", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        .MoveFirst
        If IsNull(Rs!IdVenta) = False Then
            IdTransaccion = Rs!IdVenta
        Else
            IdTransaccion = 1
        End If

        With Text1(0)
            .Text = "V-" & IdTransaccion
        End With
    End With

    For i = 1 To 3
        With Text1(i)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    For i = 0 To 2
        With Combo1(i)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    With Label1(6)
        .Visible = True
    End With

    With Combo1(2)
        .Visible = True
    End With

    With RS1
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from HZ_PARTY where cliente = 'Si' order by 2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Filter = ""
        .Requery
        If .RecordCount <> 0 Then
            With Combo1(0)
                .Clear
            End With


            While Not .EOF
                Combo1(0).AddItem .Fields(1)
                .MoveNext
            Wend
        End If
    End With
    With Rs2
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from MTL_SYSTEM_ITEMS order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Filter = ""
        .Requery
        If .RecordCount <> 0 Then
            With Combo1(1)
                .Clear
            End With

            While Not .EOF
                Combo1(1).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                .MoveNext
            Wend
        End If

        If .RecordCount = 0 Then
            MsgBox "No hay registros existentes", vbOKOnly, "Informaci?n"
            Exit Sub
        End If
    End With

    With Combo1(2)
        .AddItem "Local"
        .AddItem "Domicilio"
    End With

    With Combo2
        .AddItem "Efectivo"
        .AddItem "Tarjeta"
        .AddItem "Puntos"
    End With

    With Rs14
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from RA_TERMINAL order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Filter = ""
        .Requery
        If .RecordCount <> 0 Then
            With Combo3
                .Clear
            End With

            While Not .EOF
                Combo3.AddItem .Fields(0)
                .MoveNext
            Wend
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Form_Load" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Combo1(0)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
                With RS1
                    .Filter = ""
                    .Requery
                    v1 = 0
                    v2 = ""
                    ListaPrecios = 0
                End With
            Else
                .BackColor = COLOR_NORMAL
                With RS1
                    .Filter = ""
                    .Requery
                    .Filter = "nombre like '" & Combo1(0) & "'"
                    .Requery
                    If .RecordCount <> 0 Then v1 = .Fields(0).Value

                    If .RecordCount <> 0 Then v2 = .Fields(1).Value

                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    Case 1
        If ListaPrecios = 0 Then
            MsgBox "Seleccionar Cliente primero", vbOKOnly, "Informaci?n"
        Else
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = COLOR_NO_ENCONTRADO
                    With Rs2
                        .Filter = ""
                        .Requery
                    End With

                    With Text1(2)
                        .Text = ""
                    End With
                Else
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(7).Value, ",", ".")
                        End If
                    End With
                End If
            End With
        End If
    Case 2
        With Combo1(2)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
                With Combo1(3)
                    .Enabled = False
                End With
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    Static cadena As String
    Select Case Index
    Case 0
        With Combo1(0)
            ' si pesionamos las teclas de las flechas sale de la rutina
            If KeyCode = vbKeyUp Then Exit Sub

            If KeyCode = vbKeyDown Then Exit Sub

            If KeyCode = vbKeyLeft Then Exit Sub

            If KeyCode = vbKeyRight Then Exit Sub

            ' verifica qu no se presion? la tecla backspace
            If KeyCode <> vbKeyBack Then
                cadena = Mid(.Text, 1, Len(.Text) - .SelLength)
            Else
                '...tecla backspace
                If cadena <> "" Then
                    cadena = Mid(cadena, 1, Len(cadena) - 1)
                End If
            End If

            For i = 0 To .ListCount - 1
                If UCase(cadena) = UCase(Mid(.List(i), 1, Len(cadena))) Then
                    .ListIndex = i
                    Exit For
                End If
            Next
            ' Seelecciona
            .SelStart = Len(cadena)
            .SelLength = Len(.Text)
            If .ListIndex = -1 Then
                ' color de fondo del combo en caso de que no hay coincidencias
                .BackColor = COLOR_NO_ENCONTRADO
                With RS1
                    .Filter = ""
                    .Requery
                    v1 = 0
                    v2 = ""
                    ListaPrecios = 0
                End With
            Else
                ' Backcolor normal cuando hay coincidencia
                .BackColor = COLOR_NORMAL
                With RS1
                    .Filter = ""
                    .Filter = "nombre like '" & Combo1(0) & "'"
                    .Requery
                    If .RecordCount <> 0 Then v1 = .Fields(0)

                    If .RecordCount <> 0 Then v2 = .Fields(1)

                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    Case 1
        If ListaPrecios = 0 Then
            MsgBox "Seleccionar Cliente primero", vbOKOnly, "Informaci?n"
        Else
            With Combo1(1)
                ' si pesionamos las teclas de las flechas sale de la rutina
                If KeyCode = vbKeyUp Then Exit Sub

                If KeyCode = vbKeyDown Then Exit Sub

                If KeyCode = vbKeyLeft Then Exit Sub

                If KeyCode = vbKeyRight Then Exit Sub

                ' verifica qu no se presion? la tecla backspace
                If KeyCode <> vbKeyBack Then
                    cadena = Mid(.Text, 1, Len(.Text) - .SelLength)
                Else
                    '...tecla backspace
                    If cadena <> "" Then
                        cadena = Mid(cadena, 1, Len(cadena) - 1)
                    End If
                End If

                For i = 0 To .ListCount - 1
                    If UCase(cadena) = UCase(Mid(.List(i), 1, Len(cadena))) Then
                        .ListIndex = i
                        Exit For
                    End If
                Next
                ' Seelecciona
                .SelStart = Len(cadena)
                .SelLength = Len(.Text)
                If .ListIndex = -1 Then
                    ' color de fondo del combo en caso de que no hay coincidencias
                    .BackColor = COLOR_NO_ENCONTRADO
                    With Rs2
                        .Filter = ""
                        .Requery
                    End With

                    With Text1(2)
                        .Text = ""
                    End With
                Else
                    ' Backcolor normal cuando hay coincidencia
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(7).Value, ",", ".")
                        End If
                    End With
                End If
            End With
        End If
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_Change(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Combo1(0)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
                With RS1
                    .Filter = ""
                    .Requery
                    v1 = 0
                    v2 = ""
                    ListaPrecios = 0
                End With
            Else
                .BackColor = COLOR_NORMAL
                With RS1
                    .Filter = ""
                    .Filter = "nombre like '" & Combo1(0) & "'"
                    .Requery
                    If .RecordCount <> 0 Then v1 = .Fields(0)

                    If .RecordCount <> 0 Then v2 = .Fields(1)

                    If .RecordCount <> 0 Then If IsNull(.Fields(15).Value) = False Then ListaPrecios = .Fields(15).Value Else ListaPrecios = 1
                End With
            End If
        End With
    Case 1
        If ListaPrecios = 0 Then
            MsgBox "Seleccionar Cliente primero", vbOKOnly, "Informaci?n"
        Else
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = COLOR_NO_ENCONTRADO
                    With Rs2
                        .Filter = ""
                        .Requery
                    End With

                    With Text1(2)
                        .Text = ""
                    End With
                Else
                    .BackColor = COLOR_NORMAL
                    InItemId = Get_ItemId(.Text)
                    With Rs2
                        .Filter = "Id = " & InItemId
                        .Requery
                    End With

                    With Text1(2)
                        If ListaPrecios = 1 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(3).Value, ",", ".")
                        End If

                        If ListaPrecios = 2 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(4).Value, ",", ".")
                        End If

                        If ListaPrecios = 3 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(5).Value, ",", ".")
                        End If

                        If ListaPrecios = 4 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(6).Value, ",", ".")
                        End If

                        If ListaPrecios = 5 Then
                            If Rs2.RecordCount <> 0 Then .Text = Replace(Rs2.Fields(7).Value, ",", ".")
                        End If
                    End With
                End If
            End With
        End If
    Case 2
        With Combo1(2)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Combo1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo2_click()
    On Error GoTo errHandler
    With Combo2
        If .Text = "Puntos" Then
            With Text2(1)
                .Enabled = False
            End With
        Else
            With Text2(1)
                .Enabled = True
            End With
        End If
        If .Text = "Efectivo" And Val(Text2(0)) > Val(Text2(1)) Then
            With Label1(17)
                .Caption = "VENTA A CREDITO"
            End With
        Else
            With Label1(17)
                .Caption = "VENTA DE CONTADO"
            End With
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Combo2_click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Text1(1)
            If .Text = "" Or Val(.Text) <= 0 Then
                MsgBox "Cantidad no v?lida", vbCritical, "Error"
                Exit Sub
            End If
        End With
        If Combo1(1) <> "" And Text1(1) <> "" And Text1(2) <> "" Then
            With Rs2
                viid = .Fields(0).Value

                If ListaPrecios = 1 And Val(Text1(1)) >= 5 And .RecordCount <> 0 Then
                    viprecio = Replace(Format(.Fields(4).Value, "0.00"), ",", ".")
                Else
                    viprecio = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                End If
            End With

            With Combo1(1)
                videscripcion = Mid(.Text, 1, 47)
            End With

            With Text1(1)
                vicantidad = Replace(Format(Val(.Text), "0.00"), ",", ".")
            End With

            With Rs2
                If Get_ItemUDM(.Fields(0).Value) <> "Servicio" And Get_ItemCategoria(.Fields(0).Value) = "Inventario" Then
                    If PcInventarios = False Then
                        If Val(Get_CantidadItem(.Fields(0).Value)) < Val(vicantidad) Then
                            MsgBox "Existencia insuficiente, no se puede agregar a la venta", vbCritical, "Advertencia"
                            Exit Sub
                        End If
                    End If
                End If
            End With
            ' 1 - 10
            c1 = 10 - Len(viid)
            For i = 1 To c1
                viid = " " & viid
            Next i
            ' 12 - 58
            c2 = 47 - Len(videscripcion)
            For i = 1 To c2
                videscripcion = videscripcion & " "
            Next i
            ' 60 - 74
            c3 = 15 - Len(vicantidad)
            For i = 1 To c3
                vicantidad = " " & vicantidad
            Next i
            ' 76 - 90
            c4 = 15 - Len(viprecio)
            For i = 1 To c4
                viprecio = " " & viprecio
            Next i

            With List1
                .AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
            End With

            With Text1(1)
                .Text = ""
            End With

            With Text1(2)
                .Text = ""
            End With

            With Combo1(1)
                .Text = ""
                .BackColor = COLOR_NO_ENCONTRADO
                .SetFocus
            End With
            listSubtotal = 0
            listIva = 0
            With List1
                For i = 0 To .ListCount - 1
                    .ListIndex = i
                    .SetFocus
                    vLstCantidad = Trim(Mid(.Text, 60, 15))
                    vLstPrecio = Trim(Mid(.Text, 76, 15))
                    viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                Next i
            End With
            listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
            listIva = Replace(Format(listIva, "0.00"), ",", ".")
            listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
            listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
            With Text1(6)
                .Text = listSubtotal
            End With

            With Text1(5)
                .Text = listIva
            End With

            With Text1(4)
                .Text = listTotal
            End With
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Error"
            With Combo1(1)
                .SetFocus
            End With
        End If
    Case 1
        listSubtotal = 0
        listIva = 0
        With List1
            intX = .ListIndex
            .RemoveItem intX
            For i = 0 To .ListCount - 1
                .ListIndex = i
                .SetFocus
                vLstCantidad = Trim(Mid(.Text, 60, 15))
                vLstPrecio = Trim(Mid(.Text, 76, 15))
                viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
            Next i
        End With
        listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
        listIva = Replace(Format(listIva, "0.00"), ",", ".")
        listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
        listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
        With Text1(6)
            .Text = listSubtotal
        End With

        With Text1(5)
            .Text = listIva
        End With

        With Text1(4)
            .Text = listTotal
        End With
    Case 2
        TipoBusquedaCliente = "Venta"
        With frmBuscadorClientes
            .Show 1
        End With
    Case 3
        With frmBuscadorPedidos
            .Show 1
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Command1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1
        With Text1(1)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Case 2
        With Text1(2)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Case 3
        With Text1(3)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Text2_Change(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Text2(2)
            .Text = Replace(Format(Val(Text2(1)) - Val(Text2(0)), "0.00"), ",", ".")
        End With
    Case 1
        With Text2(2)
            .Text = Replace(Format(Val(Text2(1)) - Val(Text2(0)), "0.00"), ",", ".")
        End With

        With Label1(17)
            If Combo2 = "Efectivo" And Val(Text2(0)) > Val(Text2(1)) Then
                .Caption = "VENTA A CREDITO"
            Else
                .Caption = "VENTA DE CONTADO"
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Text2_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Anadir_Click()
    On Error GoTo errHandler
    InTipoAltaClienteProveedor = 1
    StTipoClienteProveedor = "Cliente"
    With frmClientesNuevo
        .Caption = "A?adir nuevo Cliente"
        .Show 1
    End With
    Unload frmVentas
    Set frmVentas = Nothing

    With frmVentas
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Anadir_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub List1_DblClick()
    On Error GoTo errHandler
    With Frame1(2)
        .Visible = True
    End With

    With List1
        videscripcion = Mid(List1.Text, 1, 59)
        Text1(8).Text = Trim(Mid(.Text, 60, 15))
        Text1(9).Text = Trim(Mid(.Text, 76, 15))
        intX = .ListIndex
        .RemoveItem intX
        listSubtotal = 0
        listIva = 0
        For i = 0 To .ListCount - 1
            .ListIndex = i
            .SetFocus
            vLstCantidad = Trim(Mid(.Text, 60, 15))
            vLstPrecio = Trim(Mid(.Text, 76, 15))
            viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
            listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
            listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
        Next i
    End With
    listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
    listIva = Replace(Format(listIva, "0.00"), ",", ".")
    listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
    listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
    With Text1(6)
        .Text = listSubtotal
    End With

    With Text1(5)
        .Text = listIva
    End With

    With Text1(4)
        .Text = listTotal
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:List1_DblClick" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Command3_Click()
    On Error GoTo errHandler
    With Text1(8)
        If .Text = "" Or Val(.Text) <= 0 Then
            MsgBox "Cantidad no v?lida", vbCritical, "Error"
            Exit Sub
        End If
    End With
    With Text1(9)
        If .Text = "" Or Val(.Text) <= 0 Then
            MsgBox "Precio no v?lido", vbCritical, "Error"
            Exit Sub
        End If

        If Text1(8).Text = "" Or .Text = "" Then
            MsgBox "Llene los campos que estan en blanco", vbOKOnly, "Informaci?n"
        Else
            With Text1(8)
                vicantidad = Replace(Format(Val(Trim(.Text)), "0.00"), ",", ".")
            End With
            viprecio = Replace(Format(Val(Trim(.Text)), "0.00"), ",", ".")
            ' 60 - 74
            c3 = 15 - Len(vicantidad)
            For i = 1 To c3
                vicantidad = " " & vicantidad
            Next i
            ' 76 - 90
            c4 = 15 - Len(viprecio)
            For i = 1 To c4
                viprecio = " " & viprecio
            Next i

            With List1
                .AddItem videscripcion & vicantidad & " " & viprecio
                listSubtotal = 0
                listIva = 0
                For i = 0 To .ListCount - 1
                    .ListIndex = i
                    .SetFocus
                    vLstCantidad = Trim(Mid(.Text, 60, 15))
                    vLstPrecio = Trim(Mid(.Text, 76, 15))
                    viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                Next i
            End With
            listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
            listIva = Replace(Format(listIva, "0.00"), ",", ".")
            listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
            listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
            With Text1(6)
                .Text = listSubtotal
            End With

            With Text1(5)
                .Text = listIva
            End With

            With Text1(4)
                .Text = listTotal
            End With

            With Frame1(2)
                .Visible = False
            End With
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmModificarCantidad:Command2_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    With List1
        If .ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
            With Combo1(2)
                If .Text = "" Then
                    MsgBox "Llenar el lugar de la venta", vbOKOnly, "Advertencia"
                    Exit Sub
                End If
                'confirmacion datos cliente
                If .Text = "Domicilio" Then
                    IdCliente = v1
                    With frmClientesConfirmacionDatos
                        .Show 1
                    End With
                End If
            End With

            'Mostrar pago
            With Frame1(0)
                .Enabled = False
            End With

            With Frame2(0)
                .Visible = True
            End With

            With Archivo
                .Enabled = False
            End With

            With Combo2
                .Text = "Efectivo"
            End With

            With Text2(0)
                .Text = Text1(4)
            End With

            With Text2(1)
                .Text = ""
            End With

            With Text2(3)
                .Text = Get_ClientePuntos(v1)    'Funcion Puntos
            End With

            With Label1(17)
                .Caption = "VENTA A CREDITO"
            End With
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Advertencia"
            Exit Sub
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
    vbq = MsgBox("?Desea guardar la informaci?n?", vbQuestion + vbYesNo, "Informaci?n")
    If vbq = vbYes Then
        Credito = Get_Credito(v1)
        CreditoUsado = Get_CreditoUsado(v1)
        DiasCredito = Get_CreditoDias(v1)
        DiasCreditoUsado = Get_CreditoDiasUsado(v1)
        With Combo2
            If .Text = "Tarjeta" Then
                With Text2(1)
                    If Val(.Text) = 0 Or .Text = "" Then
                        .Text = Text2(0)
                    End If
                End With
            End If
            If .Text = "Puntos" Then
                With Text2(1)
                    If Val(Text2(3)) > (Val(Text2(0))) Then
                        .Text = Text2(0).Text
                    Else
                        .Text = Text2(3).Text
                    End If
                End With
            End If
        End With

        With Text2(0)
            If Val(.Text) > Val(Text2(1)) Then
                'credito
                If Val(CreditoUsado) + Val(.Text) - Val(Text2(1)) > Val(Credito) Then
                    MsgBox "No se puede realizar la venta, el Cliente ya supero su limite de cr?dito", vbCritical, "Error"
                    'Ocultar pago
                    With Frame1(0)
                        .Enabled = True
                    End With

                    With Frame2(0)
                        .Visible = False
                    End With

                    With Archivo
                        .Enabled = True
                    End With

                    With Combo1(2)
                        .SetFocus
                    End With

                    Exit Sub
                End If
                If DiasCreditoUsado > DiasCredito And Val(.Text) - Val(Text2(1)) <> 0 Then
                    MsgBox "No se puede realizar la venta, el Cliente ya supero sus d?as de cr?dito", vbCritical, "Error"
                    'Ocultar pago
                    With Frame1(0)
                        .Enabled = True
                    End With

                    With Frame2(0)
                        .Visible = False
                    End With

                    With Archivo
                        .Enabled = True
                    End With

                    With Combo1(2)
                        .SetFocus
                    End With

                    Exit Sub
                End If
            End If
        End With

        'Actualizar Folio
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from PO_TRANSACTION_ID_R", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            .MoveFirst
            If IsNull(Rs!IdVenta) = False Then
                IdTransaccion = Rs!IdVenta
            Else
                IdTransaccion = 1
            End If

            With Text1(0)
                .Text = "V-" & IdTransaccion
            End With
        End With

        With Combo2
            'Pago con tarjeta
            If .Text = "Tarjeta" Then
                With Text2(4)
                    If .Text = "" Then
                        MsgBox "La referencia es obligatoria en el pago con tarjeta", vbCritical, "Advertencia"
                        Exit Sub
                    End If
                End With

                With Combo3
                    If .Text = "" Then
                        MsgBox "El tipo de terminal es obligatoria en el pago con tarjeta", vbCritical, "Advertencia"
                        Exit Sub
                    End If
                End With

                With Rs11
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from RA_BANK_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                    With .Fields(1)
                        .Value = Date                                                           'Fecha
                    End With

                    With .Fields(2)
                        .Value = Text2(1)                                                       'Cantidad
                    End With

                    With .Fields(3)
                        .Value = Text1(0)                                                       'Folio
                    End With

                    With .Fields(4)
                        .Value = "No"                                                           'Cancelado
                    End With

                    With .Fields(5)
                        .Value = frmMenuInicial.Combo1.Text                                     'Caja
                    End With

                    With .Fields(6)
                        .Value = Text2(4)                                                       'referencia
                    End With

                    With .Fields(7)
                        .Value = v1                                                             'cliente
                    End With

                    With .Fields(8)
                        .Value = Combo3.Text                                                    'tipo tarjeta
                    End With

                    With .Fields(9)
                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                    End With

                    With .Fields(10)
                        .Value = StUsuario                                                      'usuario
                    End With

                    With .Fields(11)
                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                    End With

                    With .Fields(12)
                        .Value = StUsuario                                                      'usuario
                    End With
                    .Update
                    .Requery
                End With
            End If

            'Pago con puntos
            If .Text = "Puntos" Then
                With Text2(3)
                    If Val(.Text) < Val(Text2(0)) Then
                        MsgBox "Puntos insuficientes para pagar la venta", vbOKOnly, "Informacion"
                        Exit Sub
                    End If
                End With

                With Text2(1)
                    If Val(.Text) <> 0 Then
                        With Rs10
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from RA_POINT_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                            With .Fields(1)
                                .Value = v1                                                             'cliente
                            End With

                            With .Fields(2)
                                .Value = Replace(Val(Text2(1)) * -1, ",", ".")                          'cantidad
                            End With

                            With .Fields(3)
                                .Value = Text1(0)                                                       'folio
                            End With

                            With .Fields(4)
                                .Value = "No"                                                           'cancelado
                            End With

                            With .Fields(5)
                                .Value = Date                                                           'fecha
                            End With

                            With .Fields(6)
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                            End With

                            With .Fields(7)
                                .Value = StUsuario                                                      'usuario
                            End With

                            With .Fields(8)
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                            End With

                            With .Fields(9)
                                .Value = StUsuario                                                      'usuario
                            End With
                            .Update
                            .Requery
                        End With
                    End If
                End With
            End If

            'pago con efectivo
            If .Text = "Efectivo" And Val(Text2(1)) > (0) Then
                With Rs9
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from RA_CASH_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                    With .Fields(1)
                        .Value = Date                                                           'fecha
                    End With

                    With .Fields(2)
                        .Value = "Pago de venta"                                                'tipo
                    End With

                    If Val(Text2(1)) >= Val(Text2(0)) Then
                        With .Fields(3)
                            .Value = Replace(Format(Val(Text2(0)), "0.00"), ",", ".")           'cantidad
                        End With
                    Else
                        With .Fields(3)
                            .Value = Replace(Format(Val(Text2(1)), "0.00"), ",", ".")           'cantidad
                        End With
                    End If

                    With .Fields(4)
                        .Value = Text1(0)                                                       'folio
                    End With

                    With .Fields(5)
                        .Value = "No"                                                           'cancelado
                    End With

                    With .Fields(6)
                        .Value = frmMenuInicial.Combo1.Text                                     'caja
                    End With

                    With .Fields(7)
                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                    End With

                    With .Fields(8)
                        .Value = StUsuario                                                      'usuario
                    End With

                    With .Fields(9)
                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                    End With

                    With .Fields(10)
                        .Value = StUsuario                                                      'usuario
                    End With
                    .Update
                    .Requery
                End With
            End If

            'si la venta no es credito
            If Val(Text2(1)) > 0 And .Text = "Efectivo" Then
                'puntos
                ClienteMayorista = Get_ClienteMayorista(v1)
                ListaPrecios = Get_ClienteListaP(v1)
                If ClienteMayorista = "No" And ListaPrecios = 1 Then
                    With Rs10
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from RA_POINT_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                        With .Fields(1)
                            .Value = v1                                                                     'cliente
                        End With

                        If Val(Text2(1)) > (Val(Text2(0))) Then
                            With .Fields(2)
                                .Value = Replace(Round(Val(Text2(0)) * Val(PcValorPuntos), 2), ",", ".")    'cantidad
                            End With
                        Else
                            With .Fields(2)
                                .Value = Replace(Round(Val(Text2(1)) * Val(PcValorPuntos), 2), ",", ".")    'cantidad
                            End With
                        End If

                        With .Fields(3)
                            .Value = Text1(0)                                                               'folio
                        End With

                        With .Fields(4)
                            .Value = "No"                                                                   'cancelado
                        End With

                        With .Fields(5)
                            .Value = Date                                                                   'fecha
                        End With

                        With .Fields(6)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")            'creacion
                        End With

                        With .Fields(7)
                            .Value = StUsuario                                                              'usuario
                        End With

                        With .Fields(8)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")            'modificacion
                        End With

                        With .Fields(9)
                            .Value = StUsuario                                                              'usuario
                        End With
                        .Update
                        .Requery
                    End With
                End If
            End If
        End With

        With Text1(0)
            v3 = .Text                  'folio
        End With
        v6 = Date                       'fecha
        v18 = "No"                      'cancelado"
        With Text1(3)
            v19 = .Text                 'comentarios
        End With
        v20 = StTipoVentasCompras       'tipo
        With Combo1(2)
            v4 = .Text                  'LugarVenta
        End With

        With Text2(1)
            DineroRestante = Val(.Text)
        End With

        With List1
            For i = 0 To .ListCount - 1
                .ListIndex = i
                'asignar valores a campos lineas
                v8 = Trim(Mid(.Text, 1, 10))                                            'idarticulo
                v7 = Get_ItemTipo(v8)                                                   'Tipoarticulo
                v9 = Get_ItemCod(v8)                                                    'codigo articulo
                v10 = Get_ItemDesc(v8)                                                  'descripcion articulo
                v11 = Replace(Format(Val(Trim(Mid(.Text, 60, 15))), "0.00"), ",", ".")  'cantidad
                v12 = Get_ItemUDM(v8)                                                   'UDM
                v13 = Replace(Format(Val(Trim(Mid(.Text, 76, 15))), "0.00"), ",", ".")  'precio
                viva = Get_ItemIva(v8)
                vCategoria = Get_ItemCategoria(v8)
                'lote
                ControlLote = Get_ItemLote(v8)
                'si tiene control de lote
                If ControlLote = True Then
                    CantidadRestante = v11
                    vCurrentLote = "V" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdTransaccion
                    'mientras no se complete la cantidad necesaria
                    While Val(CantidadRestante) > (0)
                        'obtenemos lote mas antiguo y existencia de ese lote
                        vLote = ""
                        vLote = Get_LoteConsumo(v8)
                        vCantidadLote = Get_LoteConsumoCantidad(v8)
                        'si existe algun lote
                        If vLote <> "" Then
                            If Val(vCantidadLote) > (Val(CantidadRestante)) Then
                                If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                    With Rs5
                                        If .State = 1 Then .Close
                                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                        .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                        .Requery
                                        .AddNew
                                        With .Fields(1)
                                            .Value = v8                                                             'id
                                        End With

                                        With .Fields(2)
                                            .Value = v9                                                             'codigo
                                        End With

                                        With .Fields(3)
                                            .Value = v10                                                            'descripcion
                                        End With

                                        With .Fields(4)
                                            .Value = Date                                                           'fecha
                                        End With

                                        With .Fields(5)
                                            .Value = "Salida por venta"                                             'tipo de treansaccion
                                        End With

                                        With .Fields(7)
                                            .Value = v12                                                            'udm
                                        End With

                                        With .Fields(8)
                                            .Value = v3                                                             'folio
                                        End With

                                        With .Fields(9)
                                            .Value = v18                                                            'cancelado
                                        End With

                                        With .Fields(10)
                                            .Value = vLote                                                          'lote
                                        End With

                                        With .Fields(6)
                                            .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
                                        End With

                                        With .Fields(11)
                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                        End With

                                        With .Fields(12)
                                            .Value = StUsuario                                                      'usuario
                                        End With

                                        With .Fields(13)
                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                        End With

                                        With .Fields(14)
                                            .Value = StUsuario                                                      'usuario
                                        End With
                                        .Update
                                        .Requery
                                    End With
                                End If
                                v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")               'subtotal
                                v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                                v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
                                If DineroRestante = 0 Then
                                    v17 = "0"
                                Else
                                    If DineroRestante >= Val(v16) Then
                                        v17 = v16
                                        DineroRestante = DineroRestante - Val(v16)
                                    Else
                                        v17 = Replace(Format(DineroRestante, "0.00"), ",", ".")
                                        DineroRestante = 0
                                    End If
                                End If    'totalpagado

                                'guardar compra o venta
                                With Rs3
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                    With .Fields(1)
                                        .Value = v1                                                             'idclienteproveedor
                                    End With

                                    With .Fields(2)
                                        .Value = v2                                                             'nombre cliente proveedor
                                    End With

                                    With .Fields(3)
                                        .Value = v3                                                             'folio
                                    End With

                                    With .Fields(4)
                                        .Value = v4                                                             'LugarVenta
                                    End With

                                    With .Fields(5)
                                        .Value = v6                                                             'fecha
                                    End With

                                    With .Fields(6)
                                        .Value = v7                                                             'Tipoarticulo
                                    End With

                                    With .Fields(7)
                                        .Value = v8                                                             'idarticulo
                                    End With

                                    With .Fields(8)
                                        .Value = v9                                                             'codigo articulo
                                    End With

                                    With .Fields(9)
                                        .Value = v10                                                            'descripcion articulo
                                    End With

                                    With .Fields(10)
                                        .Value = CantidadRestante                                               'cantidad
                                    End With

                                    With .Fields(11)
                                        .Value = v12                                                            'UDM
                                    End With

                                    With .Fields(12)
                                        .Value = v13                                                            'precio
                                    End With

                                    With .Fields(13)
                                        .Value = v14                                                            'subtotal
                                    End With

                                    With .Fields(14)
                                        .Value = v15                                                            'iva
                                    End With

                                    With .Fields(15)
                                        .Value = v16                                                            'total
                                    End With

                                    With .Fields(16)
                                        .Value = v17                                                            'totalpagado
                                    End With

                                    With .Fields(17)
                                        .Value = v18                                                            'cancelado
                                    End With

                                    With .Fields(18)
                                        .Value = v19                                                            'comentarios
                                    End With

                                    With .Fields(19)
                                        .Value = v20                                                            'tipo
                                    End With

                                    With .Fields(20)
                                        .Value = Replace(Replace(v3, "V-", ""), "C-", "")                       'NUM_FOLIO
                                    End With

                                    With .Fields(21)
                                        .Value = vLote                                                          'lote
                                    End With

                                    With .Fields(22)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                    End With

                                    With .Fields(23)
                                        .Value = StUsuario                                                      'usuario
                                    End With

                                    With .Fields(24)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                    End With

                                    With .Fields(25)
                                        .Value = StUsuario                                                      'usuario
                                    End With
                                    .Update
                                    .Requery
                                    .Close
                                End With
                                CantidadRestante = "0"
                            Else
                                If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                    With Rs5
                                        If .State = 1 Then .Close
                                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                        .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                        .Requery
                                        .AddNew
                                        With .Fields(1)
                                            .Value = v8                                                             'id
                                        End With

                                        With .Fields(2)
                                            .Value = v9                                                             'codigo
                                        End With

                                        With .Fields(3)
                                            .Value = v10                                                            'descripcion
                                        End With

                                        With .Fields(4)
                                            .Value = Date                                                           'fecha
                                        End With

                                        With .Fields(5)
                                            .Value = "Salida por venta"                                             'tipo de treansaccion
                                        End With

                                        With .Fields(7)
                                            .Value = v12                                                            'udm
                                        End With

                                        With .Fields(8)
                                            .Value = v3                                                             'folio
                                        End With

                                        With .Fields(9)
                                            .Value = v18                                                            'cancelado
                                        End With

                                        With .Fields(10)
                                            .Value = vLote                                                          'lote
                                        End With

                                        With .Fields(6)
                                            .Value = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".")     'cantidad
                                        End With

                                        With .Fields(11)
                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                        End With

                                        With .Fields(12)
                                            .Value = StUsuario                                                      'usuario
                                        End With

                                        With .Fields(13)
                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                        End With

                                        With .Fields(14)
                                            .Value = StUsuario                                                      'usuario
                                        End With
                                        .Update
                                        .Requery
                                    End With
                                End If
                                v14 = Replace(Format(Val(vCantidadLote) * Val(v13), "0.00"), ",", ".")              'subtotal
                                v15 = Replace(Format(Val(vCantidadLote) * Val(v13) * Val(viva), "0.00"), ",", ".")  'iva
                                v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                        'total
                                If DineroRestante = 0 Then
                                    v17 = "0"
                                Else
                                    If DineroRestante >= Val(v16) Then
                                        v17 = v16
                                        DineroRestante = DineroRestante - Val(v16)
                                    Else
                                        v17 = Replace(Format(DineroRestante, "0.00"), ",", ".")
                                        DineroRestante = 0
                                    End If
                                End If    'totalpagado

                                'guardar compra o venta
                                With Rs3
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                    With .Fields(1)
                                        .Value = v1                                                             'idclienteproveedor
                                    End With

                                    With .Fields(2)
                                        .Value = v2                                                             'nombre cliente proveedor
                                    End With

                                    With .Fields(3)
                                        .Value = v3                                                             'folio
                                    End With

                                    With .Fields(4)
                                        .Value = v4                                                             'LugarVenta
                                    End With

                                    With .Fields(5)
                                        .Value = v6                                                             'fecha
                                    End With

                                    With .Fields(6)
                                        .Value = v7                                                             'Tipoarticulo
                                    End With

                                    With .Fields(7)
                                        .Value = v8                                                             'idarticulo
                                    End With

                                    With .Fields(8)
                                        .Value = v9                                                             'codigo articulo
                                    End With

                                    With .Fields(9)
                                        .Value = v10                                                            'descripcion articulo
                                    End With

                                    With .Fields(10)
                                        .Value = vCantidadLote                                                  'cantidad
                                    End With

                                    With .Fields(11)
                                        .Value = v12                                                            'UDM
                                    End With

                                    With .Fields(12)
                                        .Value = v13                                                            'precio
                                    End With

                                    With .Fields(13)
                                        .Value = v14                                                            'subtotal
                                    End With

                                    With .Fields(14)
                                        .Value = v15                                                            'iva
                                    End With

                                    With .Fields(15)
                                        .Value = v16                                                            'total
                                    End With

                                    With .Fields(16)
                                        .Value = v17                                                            'totalpagado
                                    End With

                                    With .Fields(17)
                                        .Value = v18                                                            'cancelado
                                    End With

                                    With .Fields(18)
                                        .Value = v19                                                            'comentarios
                                    End With

                                    With .Fields(19)
                                        .Value = v20                                                            'tipo
                                    End With

                                    With .Fields(20)
                                        .Value = Replace(Replace(v3, "V-", ""), "C-", "")                       'NUM_FOLIO
                                    End With

                                    With .Fields(21)
                                        .Value = vLote                                                          'lote
                                    End With

                                    With .Fields(22)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                    End With

                                    With .Fields(23)
                                        .Value = StUsuario                                                      'usuario
                                    End With

                                    With .Fields(24)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                    End With

                                    With .Fields(25)
                                        .Value = StUsuario                                                      'usuario
                                    End With
                                    .Update
                                    .Requery
                                    .Close
                                End With
                                CantidadRestante = Replace(Format(Val(CantidadRestante) - Val(vCantidadLote), "0.00"), ",", ".")
                            End If
                            'si no existen lotes creamos uno
                        Else
                            InLoteExiste = Get_LoteExiste(vCurrentLote, v8)
                            If InLoteExiste = 0 Then
                                With Rs12
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                    With .Fields(1)
                                        .Value = v8                                                             'idarticulo
                                    End With

                                    With .Fields(2)
                                        .Value = vCurrentLote                                                   'lote
                                    End With

                                    With .Fields(3)
                                        .Value = "Venta"                                                        'tipo"
                                    End With

                                    With .Fields(4)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                    End With

                                    With .Fields(5)
                                        .Value = StUsuario                                                      'usuario
                                    End With

                                    With .Fields(6)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                    End With

                                    With .Fields(7)
                                        .Value = StUsuario                                                      'usuario
                                    End With
                                    .Update
                                    .Requery
                                    .Close
                                End With
                            End If

                            If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                With Rs5
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                    With .Fields(1)
                                        .Value = v8                                                             'id
                                    End With

                                    With .Fields(2)
                                        .Value = v9                                                             'codigo
                                    End With

                                    With .Fields(3)
                                        .Value = v10                                                            'descripcion
                                    End With

                                    With .Fields(4)
                                        .Value = Date                                                           'fecha
                                    End With

                                    With .Fields(5)
                                        .Value = "Salida por venta"                                             'tipo de treansaccion
                                    End With

                                    With .Fields(7)
                                        .Value = v12                                                            'udm
                                    End With

                                    With .Fields(8)
                                        .Value = v3                                                             'folio
                                    End With

                                    With .Fields(9)
                                        .Value = v18                                                            'cancelado
                                    End With

                                    With .Fields(10)
                                        .Value = vCurrentLote                                                   'lote
                                    End With

                                    With .Fields(6)
                                        .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
                                    End With

                                    With .Fields(11)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                    End With

                                    With .Fields(12)
                                        .Value = StUsuario                                                      'usuario
                                    End With

                                    With .Fields(13)
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                    End With

                                    With .Fields(14)
                                        .Value = StUsuario                                                      'usuario
                                    End With
                                    .Update
                                    .Requery
                                End With
                            End If
                            v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")               'subtotal
                            v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                            'total
                            If DineroRestante = 0 Then
                                v17 = "0"
                            Else
                                If DineroRestante >= Val(v16) Then
                                    v17 = v16
                                    DineroRestante = DineroRestante - Val(v16)
                                Else
                                    v17 = Replace(Format(DineroRestante, "0.00"), ",", ".")
                                    DineroRestante = 0
                                End If
                            End If    'totalpagado

                            'guardar compra o venta
                            With Rs3
                                If .State = 1 Then .Close
                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                .Requery
                                .AddNew
                                With .Fields(1)
                                    .Value = v1                                                             'idclienteproveedor
                                End With

                                With .Fields(2)
                                    .Value = v2                                                             'nombre cliente proveedor
                                End With

                                With .Fields(3)
                                    .Value = v3                                                             'folio
                                End With

                                With .Fields(4)
                                    .Value = v4                                                             'LugarVenta
                                End With

                                With .Fields(5)
                                    .Value = v6                                                             'fecha
                                End With

                                With .Fields(6)
                                    .Value = v7                                                             'Tipoarticulo
                                End With

                                With .Fields(7)
                                    .Value = v8                                                             'idarticulo
                                End With

                                With .Fields(8)
                                    .Value = v9                                                             'codigo articulo
                                End With

                                With .Fields(9)
                                    .Value = v10                                                            'descripcion articulo
                                End With

                                With .Fields(10)
                                    .Value = CantidadRestante                                               'cantidad
                                End With

                                With .Fields(11)
                                    .Value = v12                                                            'UDM
                                End With

                                With .Fields(12)
                                    .Value = v13                                                            'precio
                                End With

                                With .Fields(13)
                                    .Value = v14                                                            'subtotal
                                End With

                                With .Fields(14)
                                    .Value = v15                                                            'iva
                                End With

                                With .Fields(15)
                                    .Value = v16                                                            'total
                                End With

                                With .Fields(16)
                                    .Value = v17                                                            'totalpagado
                                End With

                                With .Fields(17)
                                    .Value = v18                                                            'cancelado
                                End With

                                With .Fields(18)
                                    .Value = v19                                                            'comentarios
                                End With

                                With .Fields(19)
                                    .Value = v20                                                            'tipo
                                End With

                                With .Fields(20)
                                    .Value = Replace(Replace(v3, "V-", ""), "C-", "")                       'NUM_FOLIO
                                End With

                                With .Fields(21)
                                    .Value = vCurrentLote                                                   'lote
                                End With

                                With .Fields(22)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                End With

                                With .Fields(23)
                                    .Value = StUsuario                                                      'usuario
                                End With

                                With .Fields(24)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                End With

                                With .Fields(25)
                                    .Value = StUsuario                                                      'usuario
                                End With
                                .Update
                                .Requery
                                .Close
                            End With
                            CantidadRestante = "0"
                        End If
                    Wend
                Else
                    v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")                'subtotal
                    v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")    'iva
                    v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".")                'total
                    With Rs5
                        If v12 <> "Servicio" And vCategoria = "Inventario" Then
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                            With .Fields(1)
                                .Value = v8                                                             'id
                            End With

                            With .Fields(2)
                                .Value = v9                                                             'codigo
                            End With

                            With .Fields(3)
                                .Value = v10                                                            'descripcion
                            End With

                            With .Fields(4)
                                .Value = Date                                                           'fecha
                            End With

                            With .Fields(5)
                                .Value = "Salida por venta"                                             'tipo de treansaccion
                            End With

                            With .Fields(6)
                                .Value = Replace(Format(Val(v11) * -1, "0.00"), ",", ".")               'cantidad
                            End With

                            With .Fields(7)
                                .Value = v12                                                            'udm
                            End With

                            With .Fields(8)
                                .Value = v3                                                             'folio
                            End With

                            With .Fields(9)
                                .Value = v18                                                            'cancelado
                            End With

                            With .Fields(11)
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                            End With

                            With .Fields(12)
                                .Value = StUsuario                                                      'usuario
                            End With

                            With .Fields(13)
                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                            End With

                            With .Fields(14)
                                .Value = StUsuario                                                      'usuario
                            End With
                            .Update
                            .Requery
                        End If

                        If .State = 1 Then .Close
                    End With

                    If DineroRestante = 0 Then
                        v17 = "0"
                    Else
                        If DineroRestante >= Val(v16) Then
                            v17 = v16
                            DineroRestante = DineroRestante - Val(v16)
                        Else
                            v17 = Replace(Format(DineroRestante, "0.00"), ",", ".")
                            DineroRestante = 0
                        End If
                    End If    'totalpagado

                    'guardar compra o venta
                    With Rs3
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                        With .Fields(1)
                            .Value = v1                                                             'idclienteproveedor
                        End With

                        With .Fields(2)
                            .Value = v2                                                             'nombre cliente proveedor
                        End With

                        With .Fields(3)
                            .Value = v3                                                             'folio
                        End With

                        With .Fields(4)
                            .Value = v4                                                             'LugarVenta
                        End With

                        With .Fields(5)
                            .Value = v6                                                             'fecha
                        End With

                        With .Fields(6)
                            .Value = v7                                                             'Tipoarticulo
                        End With

                        With .Fields(7)
                            .Value = v8                                                             'idarticulo
                        End With

                        With .Fields(8)
                            .Value = v9                                                             'codigo articulo
                        End With

                        With .Fields(9)
                            .Value = v10                                                            'descripcion articulo
                        End With

                        With .Fields(10)
                            .Value = v11                                                            'cantidad
                        End With

                        With .Fields(11)
                            .Value = v12                                                            'UDM
                        End With

                        With .Fields(12)
                            .Value = v13                                                            'precio
                        End With

                        With .Fields(13)
                            .Value = v14                                                            'subtotal
                        End With

                        With .Fields(14)
                            .Value = v15                                                            'iva
                        End With

                        With .Fields(15)
                            .Value = v16                                                            'total
                        End With

                        With .Fields(16)
                            .Value = v17                                                            'totalpagado
                        End With

                        With .Fields(17)
                            .Value = v18                                                            'cancelado
                        End With

                        With .Fields(18)
                            .Value = v19                                                            'comentarios
                        End With

                        With .Fields(19)
                            .Value = v20                                                            'tipo
                        End With

                        With .Fields(20)
                            .Value = Replace(Replace(v3, "V-", ""), "C-", "")                       'NUM_FOLIO
                        End With

                        With .Fields(22)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                        End With

                        With .Fields(23)
                            .Value = StUsuario                                                      'usuario
                        End With

                        With .Fields(24)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                        End With

                        With .Fields(25)
                            .Value = StUsuario                                                      'usuario
                        End With
                        .Update
                        .Requery
                        .Close
                    End With
                End If
            Next i
        End With

        'imprimir ticket
        With Rs6
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from PO_TRANSACTION_TICKET where folio = '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount <> 0 Then
                vTicketSubtotal = Get_SumSubtotal(.Fields(6))
                vTicketIva = Get_SumIva(.Fields(6))
                vTicketTotal = Get_SumTotal(.Fields(6))
                vTicketSubtotal = Replace(Format(Val(vTicketSubtotal), "0.00"), ",", ".")
                vTicketIva = Replace(Format(Val(vTicketIva), "0.00"), ",", ".")
                vTicketTotal = Replace(Format(Val(vTicketTotal), "0.00"), ",", ".")
                Unload dsrComprasVentas
                With dsrComprasVentas
                    Set .DataSource = Rs6

                    With .Sections("Secci?n4")
                        With .Controls("Etiqueta2")
                            .Caption = "TICKET DE VENTA"
                        End With

                        With .Controls("Etiqueta30")
                            .Caption = "Usuario: " & StUsuario
                        End With

                        With .Controls("Etiqueta3")
                            .Caption = PcNombreEmpresa
                        End With

                        With .Controls("Etiqueta4")
                            .Caption = PcRFC
                        End With

                        With .Controls("Etiqueta5")
                            .Caption = PcDireccion
                        End With

                        With .Controls("Etiqueta6")
                            .Caption = PcTelefono
                        End With

                        With .Controls("Etiqueta11")
                            .Caption = Rs6.Fields(2)    'cliente
                        End With

                        With .Controls("Etiqueta12")
                            .Caption = Rs6.Fields(3)    'calle
                        End With

                        With .Controls("Etiqueta13")
                            .Caption = Rs6.Fields(4)    'colonia
                        End With

                        With .Controls("Etiqueta14")
                            .Caption = Rs6.Fields(5)    'telefono
                        End With

                        With .Controls("Etiqueta17")
                            .Caption = Rs6.Fields(7)    'fecha
                        End With

                        With .Controls("Etiqueta18")
                            .Caption = Rs6.Fields(6)    'folio
                        End With
                    End With

                    With .Sections("Secci?n1")
                        With .Controls("Texto1")
                            .DataField = "cantidad"
                        End With

                        With .Controls("Texto2")
                            .DataField = "articulo"
                        End With

                        With .Controls("Texto3")
                            .DataField = "subtotal"
                        End With
                    End With

                    With .Sections("Secci?n5")
                        With .Controls("Etiqueta23")
                            .Caption = "$ " & vTicketSubtotal   'subtotal
                        End With

                        With .Controls("Etiqueta26")
                            .Caption = "$ " & vTicketIva        'iva
                        End With

                        With .Controls("Etiqueta27")
                            .Caption = "$ " & vTicketTotal      'total
                        End With

                        With .Controls("Label3")
                            .Caption = Get_PuntosPorVenta(Text1(0).Text)
                        End With

                        With .Controls("Label4")
                            .Caption = Text2(3).Text
                        End With

                        With .Controls("Label6")
                            .Caption = Get_Monedero(v1)
                        End With

                        If frmVentas.Combo2.Text = "Efectivo" And Val(frmVentas.Text2(1)) < Val(frmVentas.Text2(0)) Then
                            With .Controls("Etiqueta25")
                                .Visible = True
                            End With

                            With .Controls("Etiqueta28")
                                .Visible = True
                            End With
                        Else
                            With .Controls("Etiqueta25")
                                .Visible = False
                            End With

                            With .Controls("Etiqueta28")
                                .Visible = False
                            End With
                        End If
                    End With
                    .Show 1
                End With
            End If
            .Close
        End With

        With Text1(7)
            If .Text <> "" Then
                sql = "UPDATE PO_LINES_ALL SET cancelado= 'Si', last_update_date = '" & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS") & "', last_updated_by = '" & StUsuario & "' WHERE folio = '" & .Text & "'"
                With Cn
                    .Execute sql
                End With
            End If
        End With
        Unload frmVentas
        Set frmVentas = Nothing

        With frmVentas
            .Show 1
        End With
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Command2_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Salir_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Salir_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Unload dsrComprasVentas
    With Rs
        If .State = 1 Then .Close
    End With

    With RS1
        If .State = 1 Then .Close
    End With

    With Rs2
        If .State = 1 Then .Close
    End With

    With Rs3
        If .State = 1 Then .Close
    End With

    With Rs4
        If .State = 1 Then .Close
    End With

    With Rs5
        If .State = 1 Then .Close
    End With

    With Rs6
        If .State = 1 Then .Close
    End With

    With Rs9
        If .State = 1 Then .Close
    End With

    With Rs10
        If .State = 1 Then .Close
    End With

    With Rs11
        If .State = 1 Then .Close
    End With

    With Rs12
        If .State = 1 Then .Close
    End With

    With Rs13
        If .State = 1 Then .Close
    End With

    With Rs14
        If .State = 1 Then .Close
    End With

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set RS1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Rs4 = Nothing
    Set Rs5 = Nothing
    Set Rs6 = Nothing
    Set Rs9 = Nothing
    Set Rs10 = Nothing
    Set Rs11 = Nothing
    Set Rs12 = Nothing
    Set Rs13 = Nothing
    Set Rs14 = Nothing
    Set Cn = Nothing
    Set frmVentas = Nothing
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmVentas:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
