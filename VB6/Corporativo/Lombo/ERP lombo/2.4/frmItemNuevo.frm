VERSION 5.00
Begin VB.Form frmItemNuevo 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Añadir Nuevo Articulo"
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17175
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8655
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   16935
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
            Index           =   3
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   5520
            Width           =   5000
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3960
            TabIndex        =   16
            Top             =   5120
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
            Index           =   11
            Left            =   14280
            TabIndex        =   12
            Top             =   3120
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
            Index           =   10
            Left            =   14280
            TabIndex        =   10
            Top             =   2640
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
            Index           =   9
            Left            =   14280
            TabIndex        =   8
            Top             =   2160
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
            Index           =   8
            Left            =   14280
            TabIndex        =   6
            Top             =   1680
            Width           =   2500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3960
            TabIndex        =   11
            Top             =   3120
            Width           =   2504
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3960
            TabIndex        =   9
            Top             =   2640
            Width           =   2504
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3960
            TabIndex        =   7
            Top             =   2160
            Width           =   2504
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   3960
            TabIndex        =   5
            Top             =   1680
            Width           =   2504
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
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   4640
            Width           =   5000
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3960
            TabIndex        =   3
            Top             =   1200
            Width           =   2504
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
            Index           =   2
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   4120
            Width           =   5000
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
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
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   3600
            Width           =   5000
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
            Index           =   2
            Left            =   14280
            TabIndex        =   4
            Top             =   1200
            Width           =   2500
         End
         Begin VB.TextBox Text1 
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
            Height          =   420
            Index           =   1
            Left            =   3960
            TabIndex        =   2
            Top             =   720
            Width           =   12855
         End
         Begin VB.TextBox Text1 
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
            Height          =   420
            Index           =   0
            Left            =   3960
            TabIndex        =   1
            Top             =   240
            Width           =   5000
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
            Index           =   16
            Left            =   1080
            TabIndex        =   35
            Top             =   5520
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LOTEO"
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
            Index           =   15
            Left            =   1080
            TabIndex        =   34
            Top             =   5120
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 2 SIN IVA"
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
            Index           =   14
            Left            =   11400
            TabIndex        =   33
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 3 SIN IVA"
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
            Index           =   13
            Left            =   11400
            TabIndex        =   32
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 4 SIN IVA"
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
            Index           =   12
            Left            =   11400
            TabIndex        =   31
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO DE COMPRA SIN IVA"
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
            Index           =   11
            Left            =   10320
            TabIndex        =   30
            Top             =   3120
            Width           =   3735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 2 CON IVA"
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
            Index           =   10
            Left            =   1080
            TabIndex        =   29
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 3 CON IVA"
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
            Index           =   9
            Left            =   1080
            TabIndex        =   28
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 4 CON IVA"
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
            Index           =   8
            Left            =   1080
            TabIndex        =   27
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO DE COMPRA CON IVA"
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
            Index           =   7
            Left            =   -120
            TabIndex        =   26
            Top             =   3120
            Width           =   3855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CATEGORIA"
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
            Left            =   1080
            TabIndex        =   25
            Top             =   4640
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 1 CON IVA"
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
            Left            =   1080
            TabIndex        =   24
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UDM"
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
            Left            =   1080
            TabIndex        =   23
            Top             =   4120
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TASA DE IVA"
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
            Left            =   1080
            TabIndex        =   22
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO 1 SIN IVA"
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
            Left            =   11400
            TabIndex        =   21
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION"
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
            Left            =   1080
            TabIndex        =   20
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CODIGO"
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
            Left            =   1080
            TabIndex        =   19
            Top             =   240
            Width           =   2655
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
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
Attribute VB_Name = "frmItemNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
MEMEDE    '***********************************************************************************
'Nombre:        frmItemNuevo
'Proposito:     Registro de Artículos
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Se agrego confirmacion de salida sin
'                                               guardar datos
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset
Dim RS1 As New adodb.Recordset
'//OTROS
Dim i As Long
Dim In1 As Long

Private Sub Form_Load()
    On Error GoTo errHandler
    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With Text1(0)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Text1(1)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Text1(3)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Text1(4)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Text1(5)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Text1(6)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    With Text1(7)
        .BackColor = COLOR_NO_ENCONTRADO
    End With

    For i = 0 To 3
        With Combo1(i)
            .BackColor = COLOR_NORMAL
        End With
    Next i

    With RS1
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select Categoria From MTL_ITEM_CATEGORIES order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        .MoveFirst
        While Not .EOF
            Combo1(0).AddItem .Fields(0).Value
            .MoveNext
        Wend
        .MoveFirst
        Combo1(0).Text = .Fields(0).Value
        .Close
    End With

    With Combo1(1)
        .AddItem "0"
        .AddItem "0.16"
        .Text = "0"
    End With

    With Combo1(2)
        .AddItem "Kilogramo"
        .AddItem "Litro"
        .AddItem "Pieza"
        .AddItem "Servicio"
        .Text = "Kilogramo"
    End With

    With Combo1(3)
        .AddItem "Inventario"
        .AddItem "Gasto"
        .Text = "Inventario"
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Form_Load" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Combo1(2)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Case 1
        With Combo1(1)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
            Text1(2) = Replace(Format(Val(Text1(3)) / (1 + Val(.Text)), "0.00"), ",", ".")
            Text1(8) = Replace(Format(Val(Text1(4)) / (1 + Val(.Text)), "0.00"), ",", ".")
            Text1(9) = Replace(Format(Val(Text1(5)) / (1 + Val(.Text)), "0.00"), ",", ".")
            Text1(10) = Replace(Format(Val(Text1(6)) / (1 + Val(.Text)), "0.00"), ",", ".")
            Text1(11) = Replace(Format(Val(Text1(7)) / (1 + Val(.Text)), "0.00"), ",", ".")
        End With
    Case 2
        With Combo1(2)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Case 3
        With Combo1(3)
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Text1(0)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Case 1
        With Text1(1)
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

        With Text1(2)
            .Text = Replace(Format(Val(Text1(3)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        End With
    Case 4
        With Text1(4)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With

        With Text1(8)
            .Text = Replace(Format(Val(Text1(4)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        End With
    Case 5
        With Text1(5)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With

        With Text1(9)
            .Text = Replace(Format(Val(Text1(5)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        End With
    Case 6
        With Text1(6)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With

        With Text1(10)
            .Text = Replace(Format(Val(Text1(6)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        End With
    Case 7
        With Text1(7)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With

        With Text1(11)
            .Text = Replace(Format(Val(Text1(7)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" And Combo1(3) <> "" Then
            With Rs
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select count(*) as existe from MTL_SYSTEM_ITEMS where codigo like '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                In1 = .Fields(0).Value
                .Close
                If In1 = 0 Then
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from MTL_SYSTEM_ITEMS where descripcion like '" & Text1(1) & "' and UDM like '" & Combo1(2) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    In1 = .Fields(0).Value
                    .Close
                    If In1 = 0 Then
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from MTL_SYSTEM_ITEMS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                        With .Fields(1)
                            .Value = Text1(0)                                                       'Codigo
                        End With

                        With .Fields(2)
                            .Value = Text1(1)                                                       'Descripcion
                        End With

                        With .Fields(3)
                            .Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")               'Precio1
                        End With

                        With .Fields(4)
                            .Value = Replace(Format(Val(Text1(8)), "0.00"), ",", ".")               'Precio2
                        End With

                        With .Fields(5)
                            .Value = Replace(Format(Val(Text1(9)), "0.00"), ",", ".")               'Precio3
                        End With

                        With .Fields(6)
                            .Value = Replace(Format(Val(Text1(10)), "0.00"), ",", ".")              'Precio4
                        End With

                        With .Fields(7)
                            .Value = Replace(Format(Val(Text1(11)), "0.00"), ",", ".")              'Precio5
                        End With

                        With .Fields(8)
                            .Value = Combo1(1)                                                      'iva
                        End With

                        With .Fields(9)
                            .Value = Combo1(2)                                                      'udm
                        End With

                        With .Fields(10)
                            .Value = Combo1(0)                                                      'tipo
                        End With

                        With .Fields(11)
                            .Value = Check1                                                         'lote
                        End With

                        With .Fields(12)
                            .Value = Combo1(3)                                                      'categoria
                        End With

                        With .Fields(13)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                        End With

                        With .Fields(14)
                            .Value = StUsuario                                                      'usuario
                        End With

                        With .Fields(15)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                        End With

                        With .Fields(16)
                            .Value = StUsuario                                                      'usuario
                        End With
                        .Update
                        .Requery
                        .Close
                        Unload frmItemNuevo
                        Set frmItemNuevo = Nothing

                        With frmItemNuevo
                            .Show
                        End With

                        Exit Sub
                    Else
                        MsgBox "La combinación Descripción - UDM ya existe", vbCritical, "Error"
                    End If
                Else
                    MsgBox "El código esta siendo utilizado por otro artículo", vbCritical, "Error"
                End If
            End With
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Error"
        End If
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Salir_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" And Combo1(3) <> "" Then
            With Rs
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select count(*) as existe from MTL_SYSTEM_ITEMS where codigo like '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                In1 = .Fields(0).Value
                .Close
                If In1 = 0 Then
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from MTL_SYSTEM_ITEMS where descripcion like '" & Text1(1) & "' and UDM like '" & Combo1(2) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    In1 = .Fields(0).Value
                    .Close
                    If In1 = 0 Then
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from MTL_SYSTEM_ITEMS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                        With .Fields(1)
                            .Value = Text1(0)                                                       'Codigo
                        End With

                        With .Fields(2)
                            .Value = Text1(1)                                                       'Descripcion
                        End With

                        With .Fields(3)
                            .Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")               'Precio1
                        End With

                        With .Fields(4)
                            .Value = Replace(Format(Val(Text1(8)), "0.00"), ",", ".")               'Precio2
                        End With

                        With .Fields(5)
                            .Value = Replace(Format(Val(Text1(9)), "0.00"), ",", ".")               'Precio3
                        End With

                        With .Fields(6)
                            .Value = Replace(Format(Val(Text1(10)), "0.00"), ",", ".")              'Precio4
                        End With

                        With .Fields(7)
                            .Value = Replace(Format(Val(Text1(11)), "0.00"), ",", ".")              'Precio5
                        End With

                        With .Fields(8)
                            .Value = Combo1(1)                                                      'iva
                        End With

                        With .Fields(9)
                            .Value = Combo1(2)                                                      'udm
                        End With

                        With .Fields(10)
                            .Value = Combo1(0)                                                      'tipo
                        End With

                        With .Fields(11)
                            .Value = Check1                                                         'lote
                        End With

                        With .Fields(12)
                            .Value = Combo1(3)                                                      'categoria
                        End With

                        With .Fields(13)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                        End With

                        With .Fields(14)
                            .Value = StUsuario                                                      'usuario
                        End With

                        With .Fields(15)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                        End With

                        With .Fields(16)
                            .Value = StUsuario                                                      'usuario
                        End With
                        .Update
                        .Requery
                        .Close
                        Unload frmItemNuevo
                        Set frmItemNuevo = Nothing

                        Exit Sub
                    Else
                        MsgBox "La combinación Descripción - UDM ya existe", vbCritical, "Error"
                    End If
                Else
                    MsgBox "El código esta siendo utilizado por otro artículo", vbCritical, "Error"
                End If
            End With
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Error"
        End If
    Else
        Unload Me
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Salir_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    With Rs
        If .State = 1 Then .Close
    End With

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set Cn = Nothing
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
