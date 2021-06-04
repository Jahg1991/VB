VERSION 5.00
Begin VB.Form frmVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Nuevo Movimiento"
   ClientHeight    =   8310
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13815
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
   ScaleHeight     =   8310
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6480
      Index           =   0
      Left            =   3300
      TabIndex        =   0
      Top             =   915
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   6975
         Begin VB.ComboBox Combo3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   4
            Left            =   1680
            TabIndex        =   17
            Top             =   1440
            Width           =   4935
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   4935
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
            Height          =   540
            Index           =   3
            Left            =   1680
            TabIndex        =   18
            Top             =   2160
            Width           =   4935
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
            Height          =   540
            Index           =   0
            Left            =   1680
            TabIndex        =   19
            Top             =   2880
            Width           =   4935
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   1
            Left            =   1680
            TabIndex        =   20
            Top             =   3600
            Width           =   4935
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
            Height          =   540
            Index           =   2
            Left            =   1680
            TabIndex        =   21
            Top             =   4320
            Width           =   4935
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   2760
            Picture         =   "Ventas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   5520
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Terminal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   21
            Left            =   -480
            TabIndex        =   54
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   17
            Left            =   360
            TabIndex        =   43
            Top             =   5040
            Width           =   6255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   16
            Left            =   0
            TabIndex        =   42
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   15
            Left            =   -600
            TabIndex        =   41
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   0
            TabIndex        =   40
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   -480
            TabIndex        =   26
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pagado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   -600
            TabIndex        =   25
            Top             =   3720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   13
            Left            =   -480
            TabIndex        =   24
            Top             =   4440
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   2295
      Index           =   2
      Left            =   4477
      TabIndex        =   46
      Top             =   3000
      Visible         =   0   'False
      Width           =   4860
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   4575
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   420
            Index           =   8
            Left            =   1560
            TabIndex        =   48
            Top             =   120
            Width           =   2895
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   1680
            Picture         =   "Ventas.frx":080F
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   420
            Index           =   9
            Left            =   1560
            TabIndex        =   49
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7815
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   13335
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   7
            Left            =   1800
            TabIndex        =   44
            Top             =   120
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   495
            Index           =   3
            Left            =   4920
            Picture         =   "Ventas.frx":101E
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   495
            Index           =   2
            Left            =   11640
            Picture         =   "Ventas.frx":1BB9
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2640
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   0
            Left            =   1800
            TabIndex        =   3
            Top             =   720
            Width           =   9615
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   10080
            TabIndex        =   2
            Top             =   120
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   1800
            TabIndex        =   5
            Top             =   1200
            Width           =   9615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   6
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   6000
            MaxLength       =   7
            TabIndex        =   7
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "Ventas.frx":242F
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   1800
            Picture         =   "Ventas.frx":2C63
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3120
            Width           =   1455
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2025
            Left            =   240
            TabIndex        =   12
            Top             =   4080
            Width           =   12855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   3
            Left            =   1800
            TabIndex        =   8
            Top             =   2160
            Width           =   11295
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   420
            Index           =   5
            Left            =   6600
            TabIndex        =   14
            Top             =   6720
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   420
            Index           =   6
            Left            =   6600
            TabIndex        =   13
            Top             =   6240
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   6600
            TabIndex        =   15
            Top             =   7200
            Width           =   6495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   18
            Left            =   720
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Folio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   14
            Left            =   9000
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   -360
            TabIndex        =   38
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Artículo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   -360
            TabIndex        =   37
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   -360
            TabIndex        =   36
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   3360
            TabIndex        =   35
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   -360
            TabIndex        =   34
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   4440
            TabIndex        =   33
            Top             =   7200
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   -360
            TabIndex        =   32
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Iva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   4440
            TabIndex        =   31
            Top             =   6720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   10
            Left            =   4440
            TabIndex        =   30
            Top             =   6240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"Ventas.frx":3534
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   29
            Top             =   3720
            Width           =   12615
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Anadir 
         Caption         =   "Añadir Cliente"
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
    Option Explicit
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset  'folio
    Dim Rs1                 As New adodb.Recordset  'clientesproveedores
    Dim Rs2                 As New adodb.Recordset  'items
    Dim Rs3                 As New adodb.Recordset  'ventascompras
    Dim Rs4                 As New adodb.Recordset  'lista de ingredientes
    Dim Rs5                 As New adodb.Recordset  'movimientos de inventarios
    Dim Rs6                 As New adodb.Recordset  'ticket
    Dim Rs9                 As New adodb.Recordset  'movimientos de caja
    Dim Rs10                As New adodb.Recordset  'puntos
    Dim Rs11                As New adodb.Recordset  'terjeta
    Dim Rs12                As New adodb.Recordset  'lotes
    Dim Rs13                As New adodb.Recordset  'cabecera
    Dim Rs14                As New adodb.Recordset  'tipo de terminal
    
    '//OTROS
    Dim TipoErr             As Long
    Dim i                   As Long
    Dim X                   As Long
    Dim intX                As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim c3                  As Long
    Dim c4                  As Long
    Dim Prt                 As Printer
    
    '//VALORES PARA INSERTAR
    Dim v1                  As Long                 'idclienteproveedor
    Dim v2                  As String               'nombre cliente proveedor
    Dim v3                  As String               'folio
    Dim v4                  As String               'LugarVenta
    Dim v5                  As String               'mesa
    Dim v6                  As Date                 'fecha
    Dim v7                  As String               'Tipoarticulo
    Dim v8                  As Long                 'idarticulo
    Dim v9                  As String               'codigo articulo
    Dim v10                 As String               'descripcion articulo
    Dim v11                 As String               'cantidad
    Dim v12                 As String               'UDM
    Dim v13                 As String               'precio
    Dim v14                 As String               'subtotal
    Dim v15                 As String               'iva
    Dim v16                 As String               'total
    Dim v17                 As String               'totalpagado
    Dim v18                 As String               'cancelado
    Dim v19                 As String               'comentarios
    Dim v20                 As String               'tipo
    Dim IdTransaccion       As Long                 'folio
    
    '//ARTICULOS
    Dim InItemId            As Long
    Dim vCategoria          As String
    
    '//CLIENTES
    Dim Credito             As String               'credito del cliente
    Dim CreditoUsado        As String               'Credito usado del cliente
    Dim DiasCredito         As Long                 'Dias de credito del cliente
    Dim DiasCreditoUsado    As Long                 'Dias de credito usados por el cliente
    Dim ClienteMayorista    As String               '¿Es cliente mayorista?
    Dim ListaPrecios        As Long                 'lista de precios cliente
    
    '//LOTE
    Dim ControlLote         As Long
    Dim CantidadRestante    As String
    Dim vLote               As String
    Dim vCantidadLote       As String
    Dim vCurrentLote        As String
    Dim InLoteExiste        As Long
    
    '//VENTAS
    Dim listSubtotal        As String
    Dim listIva             As String
    Dim listTotal           As String
    Dim vLstCantidad        As String
    Dim vLstPrecio          As String
    Dim viva                As String
    Dim viid                As String
    Dim videscripcion       As String
    Dim vicantidad          As String
    Dim viprecio            As String
    Dim Array_Comentarios() As String
    
    '//PAGOS
    Dim DineroRestante      As String
    
    '//TICKET
    Dim vTicketSubtotal     As String
    Dim vTicketIva          As String
    Dim vTicketTotal        As String
    
    '//PEDIDOS
    Dim sql                 As String

    Private Sub Form_Load()
        On Error GoTo errHandler
        
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
            
            Text1(0).Text = "V-" & IdTransaccion
        End With
        
        For i = 1 To 3
            Text1(i).BackColor = COLOR_NO_ENCONTRADO
        Next i
        
        For i = 0 To 2
            With Combo1(i)
                .BackColor = COLOR_NO_ENCONTRADO
            End With
        Next i
        
        Label1(6).Visible = True
        
        Combo1(2).Visible = True
        
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from HZ_PARTY where cliente = 'Si' order by 2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            
            If .RecordCount <> 0 Then
                Combo1(0).Clear
                
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
                Combo1(1).Clear
                
                While Not .EOF
                    Combo1(1).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                    
                    .MoveNext
                Wend
            End If
        End With
        
        Combo1(2).AddItem "Local"
        Combo1(2).AddItem "Domicilio"
        
        If Rs2.RecordCount = 0 Then
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
            
            'Unload Me
            
            Exit Sub
        End If
        
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
                Combo3.Clear
                
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
                        
                        With Rs1
                            .Filter = ""
                            .Requery
                            
                            v1 = 0
                            v2 = ""
                            ListaPrecios = 0
                        End With
                    Else
                        .BackColor = COLOR_NORMAL
                        
                        With Rs1
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
                    MsgBox "Seleccionar Cliente primero", vbOKOnly, "Información"
                Else
                    With Combo1(1)
                        If .Text = "" Then
                            .BackColor = COLOR_NO_ENCONTRADO
                            
                            With Rs2
                                .Filter = ""
                                .Requery
                            End With
                            
                            Text1(2).Text = ""
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
                        
                        Combo1(3).Enabled = False
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
                    
                    ' verifica qu no se presionó la tecla backspace
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
                        
                        With Rs1
                            .Filter = ""
                            .Requery
                            
                            v1 = 0
                            v2 = ""
                            ListaPrecios = 0
                        End With
                    Else
                        ' Backcolor normal cuando hay coincidencia
                        .BackColor = COLOR_NORMAL
                        
                        With Rs1
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
                    MsgBox "Seleccionar Cliente primero", vbOKOnly, "Información"
                Else
                    With Combo1(1)
                        ' si pesionamos las teclas de las flechas sale de la rutina
                        If KeyCode = vbKeyUp Then Exit Sub
                        If KeyCode = vbKeyDown Then Exit Sub
                        If KeyCode = vbKeyLeft Then Exit Sub
                        If KeyCode = vbKeyRight Then Exit Sub
                        
                        ' verifica qu no se presionó la tecla backspace
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
                            
                            Text1(2).Text = ""
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
                        
                        With Rs1
                            .Filter = ""
                            .Requery
                            
                            v1 = 0
                            v2 = ""
                            ListaPrecios = 0
                        End With
                    Else
                        .BackColor = COLOR_NORMAL
                        
                        With Rs1
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
                    MsgBox "Seleccionar Cliente primero", vbOKOnly, "Información"
                Else
                    With Combo1(1)
                        If .Text = "" Then
                            .BackColor = COLOR_NO_ENCONTRADO
                            
                            With Rs2
                                .Filter = ""
                                .Requery
                            End With
                            
                            Text1(2).Text = ""
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
        
        If Combo2 = "Puntos" Then
            Text2(1).Enabled = False
        Else
            Text2(1).Enabled = True
        End If
        
        If Combo2 = "Efectivo" And Val(Text2(0)) > Val(Text2(1)) Then
            Label1(17).Caption = "Venta a Crédito"
        Else
            Label1(17).Caption = "Venta de Contado"
        End If
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
                If Text1(1) = "" Or Val(Text1(1)) <= 0 Then
                    MsgBox "Cantidad no válida", vbCritical, "Error"
                    
                    Exit Sub
                End If
                
                If Combo1(1) <> "" And Text1(1) <> "" And Text1(2) <> "" Then
                    viid = Rs2.Fields(0).Value
                    videscripcion = Mid(Combo1(1), 1, 47)
                    vicantidad = Replace(Format(Val(Text1(1)), "0.00"), ",", ".")
                    
                    If ListaPrecios = 1 And Val(Text1(1)) >= 5 And Rs2.RecordCount <> 0 Then
                        viprecio = Replace(Format(Rs2.Fields(4).Value, "0.00"), ",", ".")
                    Else
                        viprecio = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                    End If
                    
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
                    
                    List1.AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                    
                    Text1(1) = ""
                    Text1(2) = ""
                    
                    With Combo1(1)
                        .Text = ""
                        .BackColor = COLOR_NO_ENCONTRADO
                        .SetFocus
                    End With
                    
                    listSubtotal = 0
                    listIva = 0
                    
                    For i = 0 To List1.ListCount - 1
                        List1.ListIndex = i
                        
                        List1.SetFocus
                        
                        vLstCantidad = Trim(Mid(List1.Text, 60, 15))
                        vLstPrecio = Trim(Mid(List1.Text, 76, 15))
                        viva = Get_ItemIva(Trim(Mid(List1.Text, 1, 10)))
                        listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                        listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                    Next i
                    
                    listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                    listIva = Replace(Format(listIva, "0.00"), ",", ".")
                    listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                    listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                    
                    Text1(6) = listSubtotal
                    Text1(5) = listIva
                    Text1(4) = listTotal
                Else
                    MsgBox "Llenar todos los campos", vbCritical, "Error"
                    
                    Combo1(1).SetFocus
                End If
                        
            Case 1
                With List1
                    intX = .ListIndex
                    
                    .RemoveItem intX
                End With
                
                listSubtotal = 0
                listIva = 0
                
                For i = 0 To List1.ListCount - 1
                    List1.ListIndex = i
                    
                    List1.SetFocus
                    
                    vLstCantidad = Trim(Mid(List1.Text, 60, 15))
                    vLstPrecio = Trim(Mid(List1.Text, 76, 15))
                    viva = Get_ItemIva(Trim(Mid(List1.Text, 1, 10)))
                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                Next i
                
                listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                listIva = Replace(Format(listIva, "0.00"), ",", ".")
                listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                
                Text1(6) = listSubtotal
                Text1(5) = listIva
                Text1(4) = listTotal
                
            Case 2
                TipoBusquedaCliente = "Venta"
                
                frmBuscadorClientes.Show 1
            
            Case 3
                frmBuscadorPedidos.Show 1
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
            Case 4
                'With Text1(4)
                '    .Text = Round(CInt(.Text / 0.5) * 0.5, 2)
                'End With
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
                Text2(2).Text = Replace(Format(Val(Text2(1)) - Val(Text2(0)), "0.00"), ",", ".")
            
            Case 1
                Text2(2).Text = Replace(Format(Val(Text2(1)) - Val(Text2(0)), "0.00"), ",", ".")
                
                If Combo2 = "Efectivo" And Val(Text2(0)) > Val(Text2(1)) Then
                    Label1(17).Caption = "Venta a Crédito"
                Else
                    Label1(17).Caption = "Venta de Contado"
                End If
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
            .Caption = "Añadir nuevo Cliente"
            .Show 1
        End With
                
        Unload frmVentas
        
        Set frmVentas = Nothing
        
        frmVentas.Show 1
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
        
        Frame1(2).Visible = True
        
        videscripcion = Mid(List1.Text, 1, 59)
        
        Text1(8).Text = Trim(Mid(List1.Text, 60, 15))
        Text1(9).Text = Trim(Mid(List1.Text, 76, 15))
        
        With List1
            intX = .ListIndex
                    
            .RemoveItem intX
        End With
                
        listSubtotal = 0
        listIva = 0
                
        For i = 0 To List1.ListCount - 1
            List1.ListIndex = i
                    
            List1.SetFocus
                    
            vLstCantidad = Trim(Mid(List1.Text, 60, 15))
            vLstPrecio = Trim(Mid(List1.Text, 76, 15))
            viva = Get_ItemIva(Trim(Mid(List1.Text, 1, 10)))
            listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
            listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
        Next i
                
        listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
        listIva = Replace(Format(listIva, "0.00"), ",", ".")
        listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
        listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                
        Text1(6) = listSubtotal
        Text1(5) = listIva
        Text1(4) = listTotal
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
            If Text1(8) = "" Or Val(Text1(8)) <= 0 Then
                MsgBox "Cantidad no válida", vbCritical, "Error"
                        
                Exit Sub
            End If
                    
            If Text1(9) = "" Or Val(Text1(9)) <= 0 Then
                MsgBox "Precio no válido", vbCritical, "Error"
                        
                Exit Sub
            End If
                
            If Text1(8).Text = "" Or Text1(9).Text = "" Then
                MsgBox "Llene los campos que estan en blanco", vbOKOnly, "Información"
            Else
                vicantidad = Replace(Format(Val(Trim(Text1(8).Text)), "0.00"), ",", ".")
                viprecio = Replace(Format(Val(Trim(Text1(9).Text)), "0.00"), ",", ".")
                                
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
                                
                List1.AddItem videscripcion & vicantidad & " " & viprecio
                    
                listSubtotal = 0
                listIva = 0
                        
                For i = 0 To List1.ListCount - 1
                    List1.ListIndex = i
                            
                    List1.SetFocus
                            
                    vLstCantidad = Trim(Mid(List1.Text, 60, 15))
                    vLstPrecio = Trim(Mid(List1.Text, 76, 15))
                    viva = Get_ItemIva(Trim(Mid(List1.Text, 1, 10)))
                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                Next i
                        
                listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                listIva = Replace(Format(listIva, "0.00"), ",", ".")
                listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                        
                Text1(6) = listSubtotal
                Text1(5) = listIva
                Text1(4) = listTotal
                    
                Frame1(2).Visible = False
            End If
            
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
        
        If List1.ListCount <> 0 And Not IsNull(v1) And v1 <> 0 Then
            If Combo1(2) = "" Then
                MsgBox "Llenar el lugar de la venta", vbOKOnly, "Advertencia"
                
                Exit Sub
            End If
            
            'confirmacion datos cliente
            If Combo1(2) = "Domicilio" Then
                IdCliente = v1
                
                frmClientesConfirmacionDatos.Show 1
            End If
            
            'Mostrar pago
            Frame1(0).Enabled = False
            
            Frame2(0).Visible = True
            
            Archivo.Enabled = False
            
            Combo2.Text = "Efectivo"
            Text2(0) = Text1(4)
            Text2(1) = ""
            Text2(3) = Get_ClientePuntos(v1) 'Funcion Puntos
            Label1(17).Caption = "Venta a crédito"
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Advertencia"
            
            Exit Sub
        End If
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
        
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
                    
        If vbq = vbYes Then
            Credito = Get_Credito(v1)
            CreditoUsado = Get_CreditoUsado(v1)
            DiasCredito = Get_CreditoDias(v1)
            DiasCreditoUsado = Get_CreditoDiasUsado(v1)
            
            If Combo2 = "Tarjeta" Then
                If Val(Text2(1)) = 0 Or Text2(1) = "" Then
                    Text2(1) = Text2(0)
                End If
            End If
            
            If Combo2 = "Puntos" Then
                If Val(Text2(3)) > (Val(Text2(0))) Then
                    Text2(1).Text = Text2(0).Text
                Else
                    Text2(1).Text = Text2(3).Text
                End If
            End If
            
            If Val(Text2(0)) > Val(Text2(1)) Then
                'credito
                If Val(CreditoUsado) + Val(Text2(0)) - Val(Text2(1)) > Val(Credito) Then
                    MsgBox "No se puede realizar la venta, el Cliente ya supero su limite de crédito", vbCritical, "Error"
                    
                    'Ocultar pago
                    Frame1(0).Enabled = True
                    
                    Frame2(0).Visible = False
                    
                    Archivo.Enabled = True
                    
                    Combo1(2).SetFocus
                    
                    Exit Sub
                End If
            
                If DiasCreditoUsado > DiasCredito And Val(Text2(0)) - Val(Text2(1)) <> 0 Then
                    MsgBox "No se puede realizar la venta, el Cliente ya supero sus días de crédito", vbCritical, "Error"
                    
                    'Ocultar pago
                    Frame1(0).Enabled = True
                    
                    Frame2(0).Visible = False
                    
                    Archivo.Enabled = True
                    
                    Combo1(2).SetFocus
                    
                    Exit Sub
                End If
            End If
            
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
                    
                Text1(0).Text = "V-" & IdTransaccion
            End With
                
            'Pago con tarjeta
            If Combo2 = "Tarjeta" Then
                If Text2(4) = "" Then
                    MsgBox "La referencia es obligatoria en el pago con tarjeta", vbCritical, "Advertencia"
                    Exit Sub
                End If
                
                If Combo3 = "" Then
                    MsgBox "El tipo de terminal es obligatoria en el pago con tarjeta", vbCritical, "Advertencia"
                    Exit Sub
                End If
                
                With Rs11
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from RA_BANK_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                        .Fields(1) = Date
                        .Fields(2) = Text2(1)
                        .Fields(3) = Text1(0)
                        .Fields(4) = "No"
                        .Fields(5) = frmMenuInicial.Combo1.Text
                        .Fields(6) = Text2(4)
                        .Fields(7) = v1
                        .Fields(8) = Combo3.Text
                    .Update
                    .Requery
                End With
            End If
            
            'Pago con puntos
            If Combo2 = "Puntos" Then
                
                If Val(Text2(1)) <> 0 Then
                    With Rs10
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from RA_POINT_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                            .Fields(1) = v1
                            .Fields(2) = Replace(Val(Text2(1)) * -1, ",", ".")
                            .Fields(3) = Text1(0)
                            .Fields(4) = "No"
                            .Fields(5) = Date
                        .Update
                        .Requery
                    End With
                End If
            End If
                
            'pago con efectivo
            If Combo2 = "Efectivo" And Val(Text2(1)) > (0) Then
                With Rs9
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from RA_CASH_TRANSACTIONS", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                        .Fields(1) = Date
                        .Fields(2) = "Pago de venta"
                        
                        If Val(Text2(1)) >= Val(Text2(0)) Then
                            .Fields(3) = Replace(Format(Val(Text2(0)), "0.00"), ",", ".")
                        Else
                            .Fields(3) = Replace(Format(Val(Text2(1)), "0.00"), ",", ".")
                        End If
                        
                        .Fields(4) = Text1(0)
                        .Fields(5) = "No"
                        .Fields(6) = frmMenuInicial.Combo1.Text
                    .Update
                    .Requery
                End With
            End If
            
            'si la venta no es credito
            If Val(Text2(1)) > 0 And Combo2 = "Efectivo" Then
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
                            .Fields(1) = v1
                            
                            If Val(Text2(1)) > (Val(Text2(0))) Then
                                .Fields(2) = Replace(Round(Val(Text2(0)) * Val(PcValorPuntos), 2), ",", ".")
                            Else
                                .Fields(2) = Replace(Round(Val(Text2(1)) * Val(PcValorPuntos), 2), ",", ".")
                            End If
                            
                            .Fields(3) = Text1(0)
                            .Fields(4) = "No"
                            .Fields(5) = Date
                        .Update
                        .Requery
                    End With
                End If
            End If
                
            v3 = Text1(0) 'folio
            v6 = Date 'fecha
            v18 = "No" 'cancelado"
            v19 = Text1(3) 'comentarios
            v20 = StTipoVentasCompras 'tipo
            v4 = Combo1(2) 'LugarVenta
            DineroRestante = Val(Text2(1))
            
            For i = 0 To List1.ListCount - 1
                List1.ListIndex = i
                
                'asignar valores a campos lineas
                v8 = Trim(Mid(List1.Text, 1, 10)) 'idarticulo
                v7 = Get_ItemTipo(v8) 'Tipoarticulo
                v9 = Get_ItemCod(v8) 'codigo articulo
                v10 = Get_ItemDesc(v8) 'descripcion articulo
                v11 = Replace(Format(Val(Trim(Mid(List1.Text, 60, 15))), "0.00"), ",", ".") 'cantidad
                v12 = Get_ItemUDM(v8) 'UDM
                v13 = Replace(Format(Val(Trim(Mid(List1.Text, 76, 15))), "0.00"), ",", ".") 'precio
                viva = Get_ItemIva(v8)
                vCategoria = Get_ItemCategoria(v8)
                
                'lote
                ControlLote = Get_ItemLote(v8)
                                
                'si tiene control de lote
                If ControlLote = 1 Then
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
                                            .Fields(1) = v8 'id
                                            .Fields(2) = v9 'codigo
                                            .Fields(3) = v10 'descripcion
                                            .Fields(4) = Date 'fecha
                                            .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                            .Fields(7) = v12 'udm
                                            .Fields(8) = v3 'folio
                                            .Fields(9) = v18 'cancelado
                                            .Fields(10) = vLote 'lote
                                            .Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".") 'cantidad
                                        .Update
                                        .Requery
                                    End With
                                End If
                                
                                v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")   'subtotal
                                v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                                v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".") 'total
                                
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
                                End If 'totalpagado
                                
                                'guardar compra o venta
                                With Rs3
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                        .Fields(1) = v1 'idclienteproveedor
                                        .Fields(2) = v2 'nombre cliente proveedor
                                        .Fields(3) = v3 'folio
                                        .Fields(4) = v4 'LugarVenta
                                        .Fields(5) = v6 'fecha
                                        .Fields(6) = v7 'Tipoarticulo
                                        .Fields(7) = v8 'idarticulo
                                        .Fields(8) = v9 'codigo articulo
                                        .Fields(9) = v10 'descripcion articulo
                                        .Fields(10) = CantidadRestante 'cantidad
                                        .Fields(11) = v12 'UDM
                                        .Fields(12) = v13 'precio
                                        .Fields(13) = v14 'subtotal
                                        .Fields(14) = v15 'iva
                                        .Fields(15) = v16 'total
                                        .Fields(16) = v17 'totalpagado
                                        .Fields(17) = v18 'cancelado
                                        .Fields(18) = v19 'comentarios
                                        .Fields(19) = v20 'tipo
                                        .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                        .Fields(21) = vLote 'lote
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
                                            .Fields(1) = v8 'id
                                            .Fields(2) = v9 'codigo
                                            .Fields(3) = v10 'descripcion
                                            .Fields(4) = Date 'fecha
                                            .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                            .Fields(7) = v12 'udm
                                            .Fields(8) = v3 'folio
                                            .Fields(9) = v18 'cancelado
                                            .Fields(10) = vLote 'lote
                                            .Fields(6) = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".") 'cantidad
                                        .Update
                                        .Requery
                                    End With
                                End If
                                
                                v14 = Replace(Format(Val(vCantidadLote) * Val(v13), "0.00"), ",", ".")   'subtotal
                                v15 = Replace(Format(Val(vCantidadLote) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                                v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".") 'total
                                
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
                                End If 'totalpagado
                                
                                'guardar compra o venta
                                With Rs3
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                        .Fields(1) = v1 'idclienteproveedor
                                        .Fields(2) = v2 'nombre cliente proveedor
                                        .Fields(3) = v3 'folio
                                        .Fields(4) = v4 'LugarVenta
                                        .Fields(5) = v6 'fecha
                                        .Fields(6) = v7 'Tipoarticulo
                                        .Fields(7) = v8 'idarticulo
                                        .Fields(8) = v9 'codigo articulo
                                        .Fields(9) = v10 'descripcion articulo
                                        .Fields(10) = vCantidadLote 'cantidad
                                        .Fields(11) = v12 'UDM
                                        .Fields(12) = v13 'precio
                                        .Fields(13) = v14 'subtotal
                                        .Fields(14) = v15 'iva
                                        .Fields(15) = v16 'total
                                        .Fields(16) = v17 'totalpagado
                                        .Fields(17) = v18 'cancelado
                                        .Fields(18) = v19 'comentarios
                                        .Fields(19) = v20 'tipo
                                        .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                        .Fields(21) = vLote 'lote
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
                                If Rs12.State = 1 Then Rs12.Close
                                Rs12.CursorLocation = adodb.CursorLocationEnum.adUseClient
                                Rs12.Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                Rs12.Requery
                                Rs12.AddNew
                                    Rs12.Fields(1) = v8 'idarticulo
                                    Rs12.Fields(2) = vCurrentLote 'lote
                                    Rs12.Fields(3) = "Venta" 'tipo"
                                Rs12.Update
                                Rs12.Requery
                                Rs12.Close
                            End If
                            
                            If v12 <> "Servicio" And vCategoria = "Inventario" Then
                                With Rs5
                                    If .State = 1 Then .Close
                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                    .Requery
                                    .AddNew
                                        .Fields(1) = v8 'id
                                        .Fields(2) = v9 'codigo
                                        .Fields(3) = v10 'descripcion
                                        .Fields(4) = Date 'fecha
                                        .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                        .Fields(7) = v12 'udm
                                        .Fields(8) = v3 'folio
                                        .Fields(9) = v18 'cancelado
                                        .Fields(10) = vCurrentLote  'lote
                                        .Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".") 'cantidad
                                    .Update
                                    .Requery
                                End With
                            End If
                            
                            v14 = Replace(Format(Val(CantidadRestante) * Val(v13), "0.00"), ",", ".")   'subtotal
                            v15 = Replace(Format(Val(CantidadRestante) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                            v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".") 'total
                            
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
                            End If 'totalpagado
                            
                            'guardar compra o venta
                            With Rs3
                                If .State = 1 Then .Close
                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                .Requery
                                .AddNew
                                    .Fields(1) = v1 'idclienteproveedor
                                    .Fields(2) = v2 'nombre cliente proveedor
                                    .Fields(3) = v3 'folio
                                    .Fields(4) = v4 'LugarVenta
                                    .Fields(5) = v6 'fecha
                                    .Fields(6) = v7 'Tipoarticulo
                                    .Fields(7) = v8 'idarticulo
                                    .Fields(8) = v9 'codigo articulo
                                    .Fields(9) = v10 'descripcion articulo
                                    .Fields(10) = CantidadRestante 'cantidad
                                    .Fields(11) = v12 'UDM
                                    .Fields(12) = v13 'precio
                                    .Fields(13) = v14 'subtotal
                                    .Fields(14) = v15 'iva
                                    .Fields(15) = v16 'total
                                    .Fields(16) = v17 'totalpagado
                                    .Fields(17) = v18 'cancelado
                                    .Fields(18) = v19 'comentarios
                                    .Fields(19) = v20 'tipo
                                    .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                                    .Fields(21) = vCurrentLote  'lote
                                .Update
                                .Requery
                                .Close
                            End With
                                
                            CantidadRestante = "0"
                        End If
                    Wend
                Else
                    v14 = Replace(Format(Val(v11) * Val(v13), "0.00"), ",", ".")   'subtotal
                    v15 = Replace(Format(Val(v11) * Val(v13) * Val(viva), "0.00"), ",", ".")   'iva
                    v16 = Replace(Format(Val(v14) + Val(v15), "0.00"), ",", ".") 'total
                    
                    If v12 <> "Servicio" And vCategoria = "Inventario" Then
                        With Rs5
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                .Fields(1) = v8 'id
                                .Fields(2) = v9 'codigo
                                .Fields(3) = v10 'descripcion
                                .Fields(4) = Date 'fecha
                                .Fields(5) = "Salida por venta" 'tipo de treansaccion
                                .Fields(6) = Replace(Format(Val(v11) * -1, "0.00"), ",", ".") 'cantidad
                                .Fields(7) = v12 'udm
                                .Fields(8) = v3 'folio
                                .Fields(9) = v18 'cancelado
                            .Update
                            .Requery
                        End With
                    End If
                    
                    If Rs5.State = 1 Then Rs5.Close
                    
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
                    End If 'totalpagado
                    
                    'guardar compra o venta
                    With Rs3
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from PO_LINES_ALL;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                            .Fields(1) = v1 'idclienteproveedor
                            .Fields(2) = v2 'nombre cliente proveedor
                            .Fields(3) = v3 'folio
                            .Fields(4) = v4 'LugarVenta
                            .Fields(5) = v6 'fecha
                            .Fields(6) = v7 'Tipoarticulo
                            .Fields(7) = v8 'idarticulo
                            .Fields(8) = v9 'codigo articulo
                            .Fields(9) = v10 'descripcion articulo
                            .Fields(10) = v11 'cantidad
                            .Fields(11) = v12 'UDM
                            .Fields(12) = v13 'precio
                            .Fields(13) = v14 'subtotal
                            .Fields(14) = v15 'iva
                            .Fields(15) = v16 'total
                            .Fields(16) = v17 'totalpagado
                            .Fields(17) = v18 'cancelado
                            .Fields(18) = v19 'comentarios
                            .Fields(19) = v20 'tipo
                            .Fields(20) = Replace(Replace(v3, "V-", ""), "C-", "") 'NUM_FOLIO
                        .Update
                        .Requery
                        .Close
                    End With
                End If
            Next i
                
            'imprimir ticket
            With Rs6
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from PO_TRANSACTION_TICKET where folio = '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                
                If .RecordCount <> 0 Then
                    Unload TicketComprasVentas
                    
                    With TicketComprasVentas
                        vTicketSubtotal = Get_SumSubtotal(Rs6.Fields(6))
                        vTicketIva = Get_SumIva(Rs6.Fields(6))
                        vTicketTotal = Get_SumTotal(Rs6.Fields(6))
                        vTicketSubtotal = Replace(Format(Val(vTicketSubtotal), "0.00"), ",", ".")
                        vTicketIva = Replace(Format(Val(vTicketIva), "0.00"), ",", ".")
                        vTicketTotal = Replace(Format(Val(vTicketTotal), "0.00"), ",", ".")
                        Set .DataSource = Rs6
                        
                        With .Sections("Sección4")
                            .Controls("Etiqueta2").Caption = "TICKET DE VENTA"
                            .Controls("Etiqueta30").Caption = "Usuario: " & StUsuario
                            .Controls("Etiqueta3").Caption = PcNombreEmpresa
                            .Controls("Etiqueta4").Caption = PcRFC
                            .Controls("Etiqueta5").Caption = PcDireccion
                            .Controls("Etiqueta6").Caption = PcTelefono
                            .Controls("Etiqueta11").Caption = Rs6.Fields(2) 'cliente
                            .Controls("Etiqueta12").Caption = Rs6.Fields(3) 'calle
                            .Controls("Etiqueta13").Caption = Rs6.Fields(4) 'colonia
                            .Controls("Etiqueta14").Caption = Rs6.Fields(5) 'telefono
                            .Controls("Etiqueta17").Caption = Rs6.Fields(7) 'fecha
                            .Controls("Etiqueta18").Caption = Rs6.Fields(6) 'folio
                        End With
                        
                        With .Sections("Sección1")
                            .Controls("Texto1").DataField = "cantidad"
                            .Controls("Texto2").DataField = "articulo"
                            .Controls("Texto3").DataField = "subtotal"
                        End With
                        
                        With .Sections("Sección5")
                            .Controls("Etiqueta23").Caption = "$ " & vTicketSubtotal 'subtotal
                            .Controls("Etiqueta26").Caption = "$ " & vTicketIva 'iva
                            .Controls("Etiqueta27").Caption = "$ " & vTicketTotal 'total
                            .Controls("Label3").Caption = Get_PuntosPorVenta(Text1(0).Text)
                            .Controls("Label4").Caption = Text2(3).Text
                            .Controls("Label6").Caption = Get_Monedero(v1)
                            
                            If frmVentas.Combo2.Text = "Efectivo" And Val(frmVentas.Text2(1)) < Val(frmVentas.Text2(0)) Then
                                .Controls("Etiqueta25").Visible = True
                                .Controls("Etiqueta28").Visible = True
                            Else
                                .Controls("Etiqueta25").Visible = False
                                .Controls("Etiqueta28").Visible = False
                            End If
                        End With
                        
                        .Show 1
                    End With
                End If
                
                .Close
            End With
            
            If Text1(7).Text <> "" Then
                sql = "UPDATE PO_LINES_ALL SET cancelado= 'Si' WHERE folio = '" & Text1(7).Text & "'"
                
                Cn.Execute sql
            End If
            
            Unload frmVentas
            
            Set frmVentas = Nothing
            
            frmVentas.Show 1
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
        
        Unload TicketComprasVentas
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Rs2.State = 1 Then Rs2.Close
        If Rs3.State = 1 Then Rs3.Close
        If Rs4.State = 1 Then Rs4.Close
        If Rs5.State = 1 Then Rs5.Close
        If Rs6.State = 1 Then Rs6.Close
        If Rs9.State = 1 Then Rs9.Close
        If Rs10.State = 1 Then Rs10.Close
        If Rs11.State = 1 Then Rs11.Close
        If Rs12.State = 1 Then Rs12.Close
        If Rs13.State = 1 Then Rs13.Close
        If Rs14.State = 1 Then Rs14.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
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
