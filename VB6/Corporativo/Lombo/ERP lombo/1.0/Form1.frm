VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Punto de venta"
   ClientHeight    =   13590
   ClientLeft      =   3090
   ClientTop       =   3225
   ClientWidth     =   28710
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans Typewriter"
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
   ScaleHeight     =   13590
   ScaleWidth      =   28710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FBuscar 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5880
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ListBox lstBuscar 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtBuscar 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame FBuscar2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1935
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   4575
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   2175
         End
      End
   End
   Begin VB.Frame FArticulos 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   10800
      TabIndex        =   139
      Top             =   4920
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Frame FArticulos2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   4575
         Left            =   120
         TabIndex        =   140
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton CbtArticulos 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   495
            Index           =   1
            Left            =   3000
            TabIndex        =   162
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton CbtArticulos 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   495
            Index           =   0
            Left            =   3000
            TabIndex        =   159
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtNArticulo 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   1
            Left            =   2520
            TabIndex        =   157
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtNArticulo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   2
            Left            =   2520
            TabIndex        =   158
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtNArticulo 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   2520
            TabIndex        =   160
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtNArticulo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   4
            Left            =   2520
            TabIndex        =   161
            Top             =   2760
            Width           =   495
         End
         Begin VB.CheckBox cbNArticulos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2520
            TabIndex        =   163
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtNArticulo 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   5
            Left            =   2520
            TabIndex        =   164
            Top             =   3960
            Width           =   495
         End
         Begin VB.TextBox txtArticulo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   5
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   155
            Top             =   3960
            Width           =   495
         End
         Begin VB.CheckBox cbArticulos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2520
            TabIndex        =   154
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtArticulo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   4
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   153
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtArticulo 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   3
            Left            =   2520
            TabIndex        =   152
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtArticulo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   2
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   151
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtArticulo 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   1
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   150
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtArticulo 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   0
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtNArticulo 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   0
            Left            =   2520
            TabIndex        =   156
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa impositiva"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   148
            Top             =   4080
            Width           =   2295
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Control de lote"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   147
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            Height          =   375
            Index           =   5
            Left            =   720
            TabIndex        =   146
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo"
            Height          =   375
            Index           =   4
            Left            =   720
            TabIndex        =   145
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UDM"
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   144
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   143
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblArticulos 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   142
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblArticulos 
            BackStyle       =   0  'Transparent
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   141
            Top             =   0
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Fatributos 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   10800
      TabIndex        =   77
      Top             =   120
      Visible         =   0   'False
      Width           =   11175
      Begin VB.Frame Fatributos2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   3855
         Left            =   120
         TabIndex        =   78
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   30
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   138
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   29
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   137
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   28
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   136
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   27
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   135
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   26
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   134
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   25
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   133
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   24
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   132
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   23
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   131
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   22
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   130
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   21
            Left            =   8880
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   129
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   20
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   128
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   19
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   127
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   18
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   126
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   17
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   125
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   16
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   96
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   15
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   95
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   14
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   94
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   13
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   93
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   12
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   92
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   11
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   91
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   10
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   90
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   9
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   89
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   8
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   88
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   7
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   87
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   6
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   86
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   5
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   85
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   4
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   84
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   3
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   83
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   2
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   82
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtAtributos 
            Appearance      =   0  'Flat
            Height          =   345
            Index           =   1
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   81
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo3"
            Height          =   375
            Index           =   29
            Left            =   120
            TabIndex        =   124
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo4"
            Height          =   375
            Index           =   28
            Left            =   120
            TabIndex        =   123
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo5"
            Height          =   375
            Index           =   27
            Left            =   120
            TabIndex        =   122
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo6"
            Height          =   375
            Index           =   26
            Left            =   120
            TabIndex        =   121
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo7"
            Height          =   375
            Index           =   25
            Left            =   120
            TabIndex        =   120
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo8"
            Height          =   375
            Index           =   24
            Left            =   120
            TabIndex        =   119
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo9"
            Height          =   375
            Index           =   23
            Left            =   120
            TabIndex        =   118
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo10"
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   117
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo29"
            Height          =   375
            Index           =   21
            Left            =   7320
            TabIndex        =   116
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo11"
            Height          =   375
            Index           =   20
            Left            =   3720
            TabIndex        =   115
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo12"
            Height          =   375
            Index           =   19
            Left            =   3720
            TabIndex        =   114
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo13"
            Height          =   375
            Index           =   18
            Left            =   3720
            TabIndex        =   113
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo14"
            Height          =   375
            Index           =   17
            Left            =   3720
            TabIndex        =   112
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo15"
            Height          =   375
            Index           =   16
            Left            =   3720
            TabIndex        =   111
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo16"
            Height          =   375
            Index           =   15
            Left            =   3720
            TabIndex        =   110
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo17"
            Height          =   375
            Index           =   14
            Left            =   3720
            TabIndex        =   109
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo18"
            Height          =   375
            Index           =   13
            Left            =   3720
            TabIndex        =   108
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo19"
            Height          =   375
            Index           =   12
            Left            =   3720
            TabIndex        =   107
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo20"
            Height          =   375
            Index           =   11
            Left            =   3720
            TabIndex        =   106
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo30"
            Height          =   375
            Index           =   10
            Left            =   7320
            TabIndex        =   105
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo21"
            Height          =   375
            Index           =   9
            Left            =   7320
            TabIndex        =   104
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo22"
            Height          =   375
            Index           =   8
            Left            =   7320
            TabIndex        =   103
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo23"
            Height          =   375
            Index           =   7
            Left            =   7320
            TabIndex        =   102
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo24"
            Height          =   375
            Index           =   6
            Left            =   7320
            TabIndex        =   101
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo25"
            Height          =   375
            Index           =   5
            Left            =   7320
            TabIndex        =   100
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo26"
            Height          =   375
            Index           =   2
            Left            =   7320
            TabIndex        =   99
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo27"
            Height          =   375
            Index           =   1
            Left            =   7320
            TabIndex        =   98
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo28"
            Height          =   375
            Index           =   0
            Left            =   7320
            TabIndex        =   97
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo1"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo2"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   79
            Top             =   480
            Width           =   1575
         End
      End
   End
   Begin VB.Frame FEjecutarReporte 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5880
      TabIndex        =   66
      Top             =   11280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtEjecutarReporte 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame FEjecutarReporte2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1935
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   4455
         Begin MSComDlg.CommonDialog CDReportes 
            Left            =   3360
            Top             =   1440
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker DTPEjecutarReporte 
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   72
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   150274049
            CurrentDate     =   43892
         End
         Begin MSComCtl2.DTPicker DTPEjecutarReporte 
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   73
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   150274049
            CurrentDate     =   43892
         End
         Begin VB.Label lblEjecutarReporte 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   71
            Top             =   1440
            Width           =   3615
         End
         Begin VB.Label lblEjecutarReporte 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   70
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label lblEjecutarReporte 
            BackStyle       =   0  'Transparent
            Caption         =   "Ejecutar reporte"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   0
            Width           =   3615
         End
      End
   End
   Begin VB.Frame FCreacionReportes 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Grupos de concurrentes"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   5880
      TabIndex        =   53
      Top             =   7800
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame FCreacionReportes2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   3135
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   4575
         Begin VB.CheckBox cbNCreacionReportes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1800
            TabIndex        =   76
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox cbCreacionReportes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1800
            TabIndex        =   75
            Top             =   1800
            Width           =   255
         End
         Begin VB.TextBox txtNCreacionReportes 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   2
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   65
            Top             =   2280
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtNCreacionReportes 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   1
            Left            =   1800
            TabIndex        =   64
            Top             =   1080
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtCreacionReportes 
            Appearance      =   0  'Flat
            Height          =   450
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtCreacionReportes 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   0
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtCreacionReportes 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   2
            Left            =   1800
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txtNCreacionReportes 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   0
            Left            =   1800
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label LCRConsulta 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta"
            Height          =   375
            Index           =   4
            Left            =   480
            TabIndex        =   62
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label LCRParametros 
            BackStyle       =   0  'Transparent
            Caption         =   "Parámetros"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   61
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label LCreacionReportes 
            BackStyle       =   0  'Transparent
            Caption         =   "Creación de reportes"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   3615
         End
         Begin VB.Label LCRNombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   375
            Index           =   6
            Left            =   720
            TabIndex        =   59
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label LCRDescripcion 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   58
            Top             =   1200
            Width           =   1695
         End
      End
   End
   Begin VB.Frame FUsuarios 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Grupos de concurrentes"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Frame Fusuarios2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   3135
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   5415
         Begin VB.TextBox txtusuarios 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox txtusuarios 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   2
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox txtAltaUsuarios 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   2
            Left            =   1800
            TabIndex        =   51
            Top             =   1560
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ListBox lstusuarios 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   120
            TabIndex        =   46
            Top             =   2640
            Width           =   4215
         End
         Begin VB.TextBox txtusuarios 
            Appearance      =   0  'Flat
            Height          =   450
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1800
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   44
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtAltaUsuarios 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   1
            Left            =   1800
            TabIndex        =   49
            Top             =   960
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtAltaUsuarios 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.Label Lusuarios 
            BackStyle       =   0  'Transparent
            Caption         =   "Caja"
            Height          =   375
            Index           =   3
            Left            =   1080
            TabIndex        =   50
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Lusuarios 
            BackStyle       =   0  'Transparent
            Caption         =   "Responsabilidades"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Label Lusuarios 
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Lusuarios 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuarios"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   0
            Width           =   3615
         End
      End
   End
   Begin VB.Frame FGruposConcurrentes 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Grupos de concurrentes"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   11280
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ListBox LGruposConcurrentes 
         Appearance      =   0  'Flat
         Height          =   570
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox TGruposConcurrentes 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame FGruposConcurrentes2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1935
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   5415
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Grupos de concurrentes"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   3615
         End
      End
   End
   Begin VB.Frame FCajas 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Cajas"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   5880
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtCajas 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   3615
      End
      Begin VB.ListBox lstCajas 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Frame FCajas2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   4575
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cajas"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.Frame FAcciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton CCerrar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CEliminar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2640
         Picture         =   "Form1.frx":2229
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CImprimir 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2040
         Picture         =   "Form1.frx":421D
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CGuardar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1440
         Picture         =   "Form1.frx":61EA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CBuscar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         Picture         =   "Form1.frx":81F3
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CNuevo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Picture         =   "Form1.frx":A1DA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Frame FMenuInicial 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5655
      Begin VB.Frame FSubmenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   5175
         Begin VB.ListBox LSubmenu 
            Appearance      =   0  'Flat
            Height          =   2460
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   4695
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   2895
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   4935
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Submenu"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   12
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   36
               Top             =   0
               Width           =   2175
            End
         End
      End
      Begin VB.ListBox LMenuInicial 
         Appearance      =   0  'Flat
         Height          =   2190
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   5175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5895
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   5415
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Menu"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   0
            Width           =   2175
         End
      End
   End
   Begin VB.Frame FInicio 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   5880
      TabIndex        =   3
      Top             =   4920
      Width           =   4815
      Begin VB.CommandButton F1CBCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   2520
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox F1TBContraseña 
         Appearance      =   0  'Flat
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox F1TBUsuario 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton F1CBEntrar 
         Caption         =   "Entrar"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Inicio de sesión"
         Height          =   2535
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   4575
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   30
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   1695
         End
      End
   End
   Begin VB.Label LCerrarSesion 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19440
      TabIndex        =   6
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label LUser 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18240
      TabIndex        =   4
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================

'F  O   R   M   U   L   A   R   I   O

'==============================================================================================
Private Sub Form_Load()
    'Acomodamos la ventana
    Form1.WindowState = 0
    Form1.Height = FInicio.Height + 1000
    Form1.Width = FInicio.Width + 1000
    'Centramos el cuadro de inicio de sesion
    FInicio.Left = 500
    FInicio.Top = 300
    'ocultamos frames y etiquetas
    LCerrarSesion.Visible = False
    LUser.Visible = False
    FMenuInicial.Visible = False
    'limpiamos cuadros de usuario y contraseña
    F1TBUsuario = ""
    F1TBContraseña = ""
    LUser.Caption = ""
    InUserId = 0
    'limpiamos todos los frames
    ClearFrames
    'Reseteamos el contador de errores
    ErrorCount = 0
    'ocultamos el Menu Acciones
    FAcciones.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'salimos del programa
    ProgramExit
End Sub




'==============================================================================================

'I  N   I   C   I   O       D   E       S   E   S   I   O   N

'==============================================================================================
Private Sub F1CBCancelar_Click()
    'salimos del programa
    ProgramExit
End Sub
Private Sub F1CBEntrar_Click()
    Dim i As Integer
    'validamos usuario y contraseña
    StUser = "SELECT MAX(user_id) as user1,                             " & _
             "       COUNT(*)     as existe                             " & _
             "FROM fnd_user                                             " & _
             "WHERE user_name = '" & F1TBUsuario & "'                   " & _
             "  AND encrypted_user_password = '" & F1TBContraseña & "'  " & _
             "  AND end_date is NULL;                                   "
    With RsUser
        If .State = 1 Then .Close
            .Open StUser, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    'Validamos que los datos sean correctos
    If RsUser.Fields("existe") = 0 Then
        'si no existe contamos numero de errores
        ErrorCount = ErrorCount + 1
        'si el numero de errores es 3 salimos del programa
        If ErrorCount = 3 Then
            MsgBox "Demasiados intentos incorrectos el programa se cerrara", vbOKOnly, "Error"
            ProgramExit
            Exit Sub
        Else
            'si es menor de 3 mostramos advertencia
            MsgBox "El usuario o contraseña son incorrectos", vbOKOnly, "Error"
            F1TBUsuario.SetFocus
            Exit Sub
        End If
    'si es correcto
    Else
        ErrorCount = 0
        'Maximizamos la ventana
        Form1.WindowState = 2
        
        '------------------------------------------------------------------------------------------
        'VARIABLES
        '------------------------------------------------------------------------------------------
        Dim AltoVentana As Integer
        Dim AnchoVentana As Integer
        
        '------------------------------------------------------------------------------------------
        'FORMULARIO
        '------------------------------------------------------------------------------------------
        'Definimos tamaño de la ventana
        AltoVentana = Form1.Height
        AnchoVentana = Form1.Width
        Form1.WindowState = 0
        Form1.Height = AltoVentana - 500
        Form1.Width = AnchoVentana - 500
        Form1.Left = 0
        Form1.Top = 0
        'Acomodamos las etiquetas de cierre de sesion y nombre de usuario
        LCerrarSesion.Top = 50
        LUser.Top = 50
        LCerrarSesion.Left = Form1.Width - LCerrarSesion.Width
        LUser.Left = LCerrarSesion.Left - LUser.Width
        'Ponemos las etiquetas con nombre de usuario y los datos de conexion
        LUser.Caption = "Bienvenido " & F1TBUsuario
        LCerrarSesion.Caption = "Cerrar Sesion"
        'Asignamos el id del usuario
        InUserId = RsUser.Fields("user1")
        'mostramos las etiquetas usuario y cierre de sesion
        LCerrarSesion.Visible = True
        LUser.Visible = True
        
        '------------------------------------------------------------------------------------------
        'MENU INICIAL
        '------------------------------------------------------------------------------------------
        'Acomodamos el menu inicial
        FMenuInicial.Top = 1200
        FMenuInicial.Left = 50
        FMenuInicial.Visible = True
        
        '------------------------------------------------------------------------------------------
        'MENU ACCIONES
        '------------------------------------------------------------------------------------------
        'Acomodamos el menu acciones
        FAcciones.Top = 50
        FAcciones.Left = 50
        FAcciones.Visible = True
        
        '------------------------------------------------------------------------------------------
        'MENU INICIO DE SESION
        '------------------------------------------------------------------------------------------
        'Ocultamos el cuadro de inicio de sesion
        FInicio.Visible = False
        
        '------------------------------------------------------------------------------------------
        'MENU INICIAL
        '------------------------------------------------------------------------------------------
        'Mostramos el menu inicial
        FMenuInicial.Visible = True
        'Llenamos el menu inicial
        StMenuInicial = "SELECT t1.responsibility_id,                       " & _
                        "       t1.description                              " & _
                        "  FROM fnd_responsibility t1,                      " & _
                        "       fnd_user_resp_groups_direct t2              " & _
                        " WHERE t1.end_date is null                         " & _
                        "   AND t1.responsibility_id = t2.responsibility_id " & _
                        "   AND t2.end_date is null                         " & _
                        "   AND t2.user_id = " & InUserId & ";              "
        With RsMenuInicial
            If .State = 1 Then .Close
                .Open StMenuInicial, Cn, adOpenStatic, adLockOptimistic
                .Requery
                .MoveFirst
        End With
        While Not RsMenuInicial.EOF
            If Not IsNull(RsMenuInicial.Fields("description")) Then
                LMenuInicial.AddItem RsMenuInicial.Fields("description")
            End If
            RsMenuInicial.MoveNext
        Wend
        RsMenuInicial.Close
        
        '------------------------------------------------------------------------------------------
        'MENU BUSQUEDA
        '------------------------------------------------------------------------------------------
        'Acomodamos el frame de busqueda
        FBuscar.Top = 3000
        FBuscar.Left = FMenuInicial.Width + 1500
        FBuscar.Width = Form1.Width / 4
        FBuscar.Height = Form1.Height / 4
        FBuscar2.Width = FBuscar.Width - 240
        FBuscar2.Height = FBuscar.Height - 240
        txtBuscar.Width = FBuscar.Width - 500
        lstBuscar.Width = FBuscar.Width - 500
        lstBuscar.Height = FBuscar.Height - 1400
        
        '------------------------------------------------------------------------------------------
        'ATRIBUTOS
        '------------------------------------------------------------------------------------------
        'Acomodamos el frame de busqueda
        Fatributos.Top = 3000
        Fatributos.Left = FMenuInicial.Width + 1500
        
        '------------------------------------------------------------------------------------------
        'ADMINISTRADOR DEL SISTEMA
        '------------------------------------------------------------------------------------------
            '======================================================================================
            'GRUPOS DE CONCURRENTES
            '======================================================================================
            'Acomodamos el frame de grupos de concurrentes
            FGruposConcurrentes.Top = 1200
            FGruposConcurrentes.Left = FMenuInicial.Width + 500
            FGruposConcurrentes.Width = Form1.Width - FGruposConcurrentes.Left - 1000
            FGruposConcurrentes.Height = Form1.Height - 2000
            FGruposConcurrentes2.Width = FGruposConcurrentes.Width - 240
            FGruposConcurrentes2.Height = FGruposConcurrentes.Height - 240
            TGruposConcurrentes.Width = FGruposConcurrentes.Width - 450
            LGruposConcurrentes.Width = FGruposConcurrentes.Width - 450
            LGruposConcurrentes.Height = FGruposConcurrentes.Height - 1450
            
            '======================================================================================
            'CAJAS
            '======================================================================================
            'Acomodamos el frame de cajas
            FCajas.Top = 1200
            FCajas.Left = FMenuInicial.Width + 500
            FCajas.Width = Form1.Width - FCajas.Left - 1000
            FCajas.Height = Form1.Height - 2000
            FCajas2.Width = FCajas.Width - 240
            FCajas2.Height = FCajas.Height - 240
            txtCajas.Width = FCajas.Width - 450
            lstCajas.Width = FCajas.Width - 450
            lstCajas.Height = FCajas.Height - 1450
            
            '======================================================================================
            'USUARIOS
            '======================================================================================
            'Acomodamos el frame de usuarios
            FUsuarios.Top = 1200
            FUsuarios.Left = FMenuInicial.Width + 500
            FUsuarios.Width = Form1.Width - FUsuarios.Left - 1000
            FUsuarios.Height = Form1.Height - 2000
            Fusuarios2.Width = FUsuarios.Width - 240
            Fusuarios2.Height = FUsuarios.Height - 240
            txtusuarios(0).Width = FUsuarios.Width - 450
            txtusuarios(1).Width = txtusuarios(0).Width - 1680
            txtusuarios(2).Width = txtusuarios(0).Width - 1680
            txtAltaUsuarios(0).Width = txtusuarios(0).Width
            txtAltaUsuarios(1).Width = txtusuarios(1).Width
            txtAltaUsuarios(2).Width = txtusuarios(1).Width
            lstusuarios.Width = FUsuarios.Width - 450
            lstusuarios.Height = FUsuarios.Height - 2900
            
            '======================================================================================
            'CREACION DE REPORTES
            '======================================================================================
            'Acomodamos el frame de creacion de reportes
            FCreacionReportes.Top = 1200
            FCreacionReportes.Left = FMenuInicial.Width + 500
            FCreacionReportes.Width = Form1.Width - FCreacionReportes.Left - 1000
            FCreacionReportes.Height = Form1.Height - 2000
            FCreacionReportes2.Width = FCreacionReportes.Width - 240
            FCreacionReportes2.Height = FCreacionReportes.Height - 240
            txtCreacionReportes(0).Width = FCreacionReportes.Width - 2300
            txtCreacionReportes(1).Width = txtCreacionReportes(0).Width
            txtCreacionReportes(2).Width = txtCreacionReportes(0).Width
            txtCreacionReportes(2).Height = FCreacionReportes.Height - 2900
            txtNCreacionReportes(0).Width = txtCreacionReportes(0).Width
            txtNCreacionReportes(1).Width = txtCreacionReportes(0).Width
            txtNCreacionReportes(2).Width = txtCreacionReportes(0).Width
            txtNCreacionReportes(2).Height = txtCreacionReportes(2).Height
            
            '======================================================================================
            'EJECUCION DE REPORTES
            '======================================================================================
            'Acomodamos el frame de ejecucion de reportes
            FEjecutarReporte.Top = 1200
            FEjecutarReporte.Left = FMenuInicial.Width + 500
            FEjecutarReporte.Width = Form1.Width - FEjecutarReporte.Left - 1000
            FEjecutarReporte.Height = Form1.Height - 2000
            FEjecutarReporte2.Width = FEjecutarReporte.Width - 240
            FEjecutarReporte2.Height = FEjecutarReporte.Height - 240
            txtEjecutarReporte.Width = FEjecutarReporte.Width - 450
            
        '------------------------------------------------------------------------------------------
        'INVENTARIO
        '------------------------------------------------------------------------------------------
            '======================================================================================
            'ARTICULOS
            '======================================================================================
            'Acomodamos el frame de articulos
            FArticulos.Top = 1200
            FArticulos.Left = FMenuInicial.Width + 500
            FArticulos.Width = Form1.Width - FCajas.Left - 1000
            FArticulos.Height = Form1.Height - 2000
            FArticulos2.Width = FArticulos.Width - 240
            FArticulos2.Height = FArticulos.Height - 240
            For i = 0 To 5
                txtArticulo(i).Width = FArticulos.Width - 3000
                txtNArticulo(i).Width = FArticulos.Width - 3615
            Next i
            For i = 0 To 1
                CbtArticulos(i).Left = txtNArticulo(i).Width + txtNArticulo(i).Left
            Next i
    End If
    'cerramos el recordset
    RsUser.Close
End Sub

Private Sub LCerrarSesion_Click()
    'reiniciamos el programa
    Unload Form1
    main
End Sub




'==============================================================================================

'M  E   N   U       I   N   I   C   I   A   L

'==============================================================================================
Private Sub LMenuInicial_Click()
    'limpiamos el submenu
    LSubmenu.Clear
    'obtenemos el InRespId
    InRespId = GetRespId(LMenuInicial.Text)
    'llenamos el submenu
    StSubMenu = "SELECT description                             " & _
                "  FROM fnd_responsibility_menu                 " & _
                " WHERE end_date is null                        " & _
                "   AND responsibility_id = " & InRespId & ";   "
    With RsSubMenu
        If .State = 1 Then .Close
            .Open StSubMenu, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    While Not RsSubMenu.EOF
        If Not IsNull(RsSubMenu.Fields("description")) Then
            LSubmenu.AddItem RsSubMenu.Fields("description")
        End If
        RsSubMenu.MoveNext
    Wend
    'cerramos el recordset
    RsSubMenu.Close
End Sub




'==============================================================================================

'S  U   B   M   E   N   U

'==============================================================================================
Private Sub LSubmenu_Click()
    On Error Resume Next
    'limpiamos los frames
    ClearFrames
    'obtenemos el frame a mostrar
    SubMenuFrameName = ""
    SubMenuFrameName = GetSubMenuFrame(LSubmenu.Text, InRespId)
    FMenuInicial.Enabled = False
    'limpiamos variables
    'StTipoBuscador = ""
    'Mostramos el frame seleccionado
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'GRUPOS DE CONCURRENTES
        '======================================================================================
        'si el frame es grupos de concurrentes
        If SubMenuFrameName = "FGruposConcurrentes" Then
            'mostramos el frame
            FGruposConcurrentes.Visible = True
            'cargamos los datos
            SubMenuSt1 = "SELECT request_group_id,  " & _
                         "       description        " & _
                         "FROM fnd_request_groups;  "
            With SubMenuRs1
                If .State = 1 Then .Close
                    .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            Set TGruposConcurrentes.DataSource = SubMenuRs1
            TGruposConcurrentes.DataField = SubMenuRs1.Fields(1).Name
            'salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'CAJAS
        '======================================================================================
        'si el frame es cajas
        If SubMenuFrameName = "FCajas" Then
            'llenamos la lista
            lstCajas.Clear
            SubMenuSt1 = "SELECT caja_id,       " & _
                         "       description    " & _
                         "FROM fnd_cajas;       "
            With SubMenuRs1
                If .State = 1 Then .Close
                    .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not SubMenuRs1.EOF
                If Not IsNull(SubMenuRs1.Fields("description")) Then
                    lstCajas.AddItem SubMenuRs1.Fields("description")
                End If
                SubMenuRs1.MoveNext
            Wend
            'salimos del procedimiento
            'mostramos el frame
            FCajas.Visible = True
            'habilitamos los botones
            CGuardar.Enabled = True
            CEliminar.Enabled = True
            Exit Sub
        End If
        
        '======================================================================================
        'USUARIOS
        '======================================================================================
        'si el frame es usuarios
        If SubMenuFrameName = "FUsuarios" Then
            'mostramos el frame
            FUsuarios.Visible = True
            'cargamos los datos
            SubMenuSt1 = "SELECT t1.user_id,                    " & _
                         "       t1.user_name,                  " & _
                         "       t2.caja_id,                    " & _
                         "       t2.description,                " & _
                         "       t1.encrypted_user_password     " & _
                         "  FROM fnd_user  t1,                  " & _
                         "       fnd_cajas t2                   " & _
                         " WHERE isnull(t1.caja,1) = t2.caja_id; "
            With SubMenuRs1
                If .State = 1 Then .Close
                    .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            Set txtusuarios(0).DataSource = SubMenuRs1
            txtusuarios(0).DataField = SubMenuRs1.Fields(1).Name
            Set txtusuarios(1).DataSource = SubMenuRs1
            txtusuarios(1).DataField = SubMenuRs1.Fields(4).Name
            Set txtusuarios(2).DataSource = SubMenuRs1
            txtusuarios(2).DataField = SubMenuRs1.Fields(3).Name
            'salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'CREACION DE REPORTES
        '======================================================================================
        'si el frame es creacion de reportes
        If SubMenuFrameName = "FCreacionReportes" Then
            'mostramos el frame
            FCreacionReportes.Visible = True
            'cargamos los datos
            SubMenuSt1 = "SELECT request_header_id  " & _
                         "   ,request_unit_name     " & _
                         "   ,description           " & _
                         "   ,parametros            " & _
                         "   ,query                 " & _
                         "FROM fnd_request_headers; "
            With SubMenuRs1
                If .State = 1 Then .Close
                    .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            Set txtCreacionReportes(0).DataSource = SubMenuRs1
            txtCreacionReportes(0).DataField = SubMenuRs1.Fields(1).Name
            Set txtCreacionReportes(1).DataSource = SubMenuRs1
            txtCreacionReportes(1).DataField = SubMenuRs1.Fields(2).Name
            Set cbCreacionReportes.DataSource = SubMenuRs1
            cbCreacionReportes.DataField = SubMenuRs1.Fields(3).Name
            Set txtCreacionReportes(2).DataSource = SubMenuRs1
            txtCreacionReportes(2).DataField = SubMenuRs1.Fields(4).Name
            'salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'EJECUCION DE REPORTES
        '======================================================================================
        'si el frame es ejecucion de reportes
        If SubMenuFrameName = "FEjecutarReporte" Then
            'mostramos el frame
            FEjecutarReporte.Visible = True
            'habilitamos los botones
            CBuscar.Enabled = True
            'salimos del procedimiento
            Exit Sub
        End If
        
    '------------------------------------------------------------------------------------------
    'INVENTARIOS
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'ARTICULOS
        '======================================================================================
        'si el frame es articulos
        If SubMenuFrameName = "FArticulos" Then
            'asignamos valor a las variables
            StTipoBuscador = "Articulos"
            StTipoGuardarArticulo = "Articulos"
            'mostramos el frame
            FArticulos.Visible = True
            'cargamos los datos
            SubMenuSt1 = "SELECT                            " & _
                         "  t1.inventory_item_id,           " & _
                         "  t1.segment1,                    " & _
                         "  t1.description,                 " & _
                         "  t1.uom,                         " & _
                         "  t1.item_cost,                   " & _
                         "  t2.description as category,     " & _
                         "  t1.lot_control,                 " & _
                         "  t1.tax_rate                     " & _
                         "FROM                              " & _
                         "  mtl_system_items_b t1,          " & _
                         "  mtl_item_categories t2          " & _
                         "WHERE                             " & _
                         "  t1.category_id = t2.category_id " & _
                         "ORDER BY 1;                       "
            With SubMenuRs1
                If .State = 1 Then .Close
                    .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            Set txtArticulo(0).DataSource = SubMenuRs1
            txtArticulo(0).DataField = SubMenuRs1.Fields(1).Name
            Set txtArticulo(1).DataSource = SubMenuRs1
            txtArticulo(1).DataField = SubMenuRs1.Fields(2).Name
            Set txtArticulo(2).DataSource = SubMenuRs1
            txtArticulo(2).DataField = SubMenuRs1.Fields(3).Name
            Set txtArticulo(3).DataSource = SubMenuRs1
            txtArticulo(3).DataField = SubMenuRs1.Fields(4).Name
            Set txtArticulo(4).DataSource = SubMenuRs1
            txtArticulo(4).DataField = SubMenuRs1.Fields(5).Name
            Set cbArticulo.DataSource = SubMenuRs1
            cbArticulo.DataField = SubMenuRs1.Fields(6).Name
            Set txtArticulo(5).DataSource = SubMenuRs1
            txtArticulo(5).DataField = SubMenuRs1.Fields(7).Name
            'habilitamos los botones
            CGuardar.Enabled = True
            CNuevo.Enabled = True
            CBuscar.Enabled = True
            StTipoBuscadorArticulo = "Articulo"
            'salimos del procedimiento
            Exit Sub
        End If
        
End Sub




'==============================================================================================

'M  E   N   U       A   C   C   I   O   N   E   S

'==============================================================================================
Private Sub CCerrar_Click()
    '------------------------------------------------------------------------------------------
    'PROGRAMA PRINCIPAL
    '------------------------------------------------------------------------------------------
    'Si no hay ningun frame visible cerramos el Programa
    If FMenuInicial.Enabled = True Then
        ProgramExit
        'Salimos del procedimiento
        Exit Sub
    End If
    
    '------------------------------------------------------------------------------------------
    'BUSCADOR
    '------------------------------------------------------------------------------------------
    'Cierre del Menu Buscar
    If FBuscar.Visible = True Then
        'Ocultamos el frame
        FBuscar.Visible = False
        'limpiamos los campos
        txtBuscar.Text = ""
        lstBuscar.Clear
        
        '======================================================================================
        'ADMINISTRADOR DEL SISTEMA
        '======================================================================================
            '..................................................................................
            'GRUPOS DE CONCURRENTES
            '..................................................................................
            'Si el menu buscar estaba siendo usado por el frame Grupos de concurrentes
            If FGruposConcurrentes.Visible = True And FGruposConcurrentes.Enabled = False Then
                'habilitamos el frame
                FGruposConcurrentes.Enabled = True
                'Habiltmos/deshabilitamos botones
                CNuevo.Enabled = True
                CGuardar.Enabled = False
                CEliminar.Enabled = True
                'Salimos del procedimiento
                Exit Sub
            End If
            
            '..................................................................................
            'USUARIOS
            '..................................................................................
            'Si el menu buscar estaba siendo usado por el frame Grupos de concurrentes
            If FUsuarios.Visible = True And FUsuarios.Enabled = False Then
                'habilitamos el frame
                FUsuarios.Enabled = True
                'Habiltmos/deshabilitamos botones
                CGuardar.Enabled = True
                CEliminar.Enabled = True
                If InUusario = 2 Then
                    CNuevo.Enabled = True
                Else
                    CNuevo.Enabled = False
                End If
                'Salimos del procedimiento
                Exit Sub
            End If
            
            '..................................................................................
            'EJECUCION DE REPORTES
            '..................................................................................
            'Si el menu buscar estaba siendo usado por el frame Grupos de concurrentes
            If FEjecutarReporte.Visible = True And FEjecutarReporte.Enabled = False Then
                'habilitamos el frame
                FEjecutarReporte.Enabled = True
                'Habiltmos/deshabilitamos botones
                CBuscar.Enabled = True
                'Salimos del procedimiento
                Exit Sub
            End If
            
        '======================================================================================
        'INVENTARIO
        '======================================================================================
            '..................................................................................
            'ARTICULOS
            '..................................................................................
            'Si el menu buscar estaba siendo usado por el frame Articulos
            If FArticulos.Visible = True And FArticulos.Enabled = False Then
                'habilitamos el frame
                FArticulos.Enabled = True
                'Habiltmos/deshabilitamos botones
                CNuevo.Enabled = True
                CGuardar.Enabled = True
                CBuscar.Enabled = True
                StTipoGuardarArticulo = "Articulos"
                'Salimos del procedimiento
                Exit Sub
            End If
    End If
    
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'GRUPOS DE CONCURRENTES
        '======================================================================================
        'Cierre del frameGrupos de concurrentes
        If FGruposConcurrentes.Visible = True And FGruposConcurrentes.Enabled = True Then
            'Ocultamos el frame
            FGruposConcurrentes.Visible = False
            'habilitamos el menu inicial
            FMenuInicial.Enabled = True
            'deshabilitamos los botones nuevo y eliminar
            CNuevo.Enabled = False
            CEliminar.Enabled = False
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'CAJAS
        '======================================================================================
        'Cierre del frame cajas
        If FCajas.Visible = True Then
            'Ocultamos el frame
            FCajas.Visible = False
            'habilitamos el menu inicial
            FMenuInicial.Enabled = True
            'deshabilitamos los botones nuevo y eliminar
            CGuardar.Enabled = False
            CEliminar.Enabled = False
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'USUARIOS
        '======================================================================================
        'Cierre del frame usuarios
        If FUsuarios.Visible = True And FUsuarios.Enabled = True Then
            'Ocultamos el frame
            FUsuarios.Visible = False
            'habilitamos el menu inicial
            FMenuInicial.Enabled = True
            'deshabilitamos los botones
                CNuevo.Enabled = False
                CGuardar.Enabled = False
                CEliminar.Enabled = False
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'CREACION DE REPORTES
        '======================================================================================
        'Cierre del frame creacion de reportes
        If FCreacionReportes.Visible = True And FCreacionReportes.Enabled = True Then
            'Ocultamos el frame
            FCreacionReportes.Visible = False
            'habilitamos el menu inicial
            FMenuInicial.Enabled = True
            'deshabilitamos los botones
            CGuardar.Enabled = False
            CEliminar.Enabled = False
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'EJECUCION DE REPORTES
        '======================================================================================
        'Cierre del frame ejecucion de reportes
        If FEjecutarReporte.Visible = True And FEjecutarReporte.Enabled = True Then
            'Ocultamos el frame
            FEjecutarReporte.Visible = False
            'habilitamos el menu inicial
            FMenuInicial.Enabled = True
            'deshabilitamos los botones
            CGuardar.Enabled = False
            CBuscar.Enabled = False
            'Salimos del procedimiento
            Exit Sub
        End If
    
    '------------------------------------------------------------------------------------------
    'INVENTARIO
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'ARTICULOS
        '======================================================================================
        'Cierre del frame articulos
        If FArticulos.Visible = True And FArticulos.Enabled = True Then
            'Ocultamos el frame
            FArticulos.Visible = False
            'habilitamos el menu inicial
            FMenuInicial.Enabled = True
            'deshabilitamos los botones nuevo y eliminar
            CNuevo.Enabled = False
            CGuardar.Enabled = False
            CBuscar.Enabled = False
            'Salimos del procedimiento
            Exit Sub
        End If
End Sub
Private Sub CEliminar_Click()
    '------------------------------------------------------------------------------------------
    'variables
    '------------------------------------------------------------------------------------------
    Dim stString As String
    Dim DeleteFromTable As New ADODB.Command
    Dim rsRecordset As New ADODB.Recordset
    Dim InInteger As Integer
    Dim InInteger2 As Integer
    
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'GRUPOS DE CONCURRENTES
        '======================================================================================
        'Si es utilizado en Grupos de concurrentes
        If FGruposConcurrentes.Visible = True Then
            On Error GoTo err
            'Si no se selecciono nada en la lista
            If LGruposConcurrentes.Text = "" Then
                MsgBox "Seleccionar algun concurrente", vbOKOnly, "Informacion"
            'Eliminamos el registro
            Else
                With DeleteFromTable
                    .CommandText = "DELETE FROM fnd_request_group_units                                         " & _
                                   " WHERE request_group_id = '" & SubMenuRs1.Fields("request_group_id") & "'   " & _
                                   "   AND description      = '" & LGruposConcurrentes.Text & "';               "
                    .ActiveConnection = Cn
                    .Execute
                End With
                'limpiamos la lista y cargamos los datos actualizados
                LGruposConcurrentes.Clear
                With SubMenuRs2
                    If .State = 1 Then .Close
                        .Open SubMenuSt2, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not SubMenuRs2.EOF
                    If Not IsNull(SubMenuRs2.Fields("description")) Then
                        LGruposConcurrentes.AddItem SubMenuRs2.Fields("description")
                    End If
                    SubMenuRs2.MoveNext
                Wend
                SubMenuRs2.Close
            End If
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'CAJAS
        '======================================================================================
        'Si es utilizado en cajas
        If FCajas.Visible = True Then
            On Error GoTo err
            'Si no se selecciono nada en la lista
            If lstCajas.Text = "" Then
                MsgBox "Seleccionar alguna caja", vbOKOnly, "Informacion"
            Else
                'validamos que no este siedo usada por ningun usuario
                InInteger2 = Getcaja_id(lstCajas.Text)
                stString = "SELECT count(*) as existe           " & _
                           "  FROM fnd_user                     " & _
                           " WHERE caja = " & InInteger2 & ";   "
                 With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                InInteger = rsRecordset.Fields("existe")
                rsRecordset.Close
                If InInteger <> 0 Then
                    MsgBox "La caja esta en uso por algun usuario", vbOKOnly, "Informacion"
                Else
                    With DeleteFromTable
                        .CommandText = "DELETE FROM fnd_cajas                   " & _
                                       " WHERE caja_id = " & InInteger2 & ";    "
                        .ActiveConnection = Cn
                        .Execute
                    End With
                    'limpiamos la lista y cargamos los datos actualizados
                    lstCajas.Clear
                    With SubMenuRs1
                        If .State = 1 Then .Close
                            .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    While Not SubMenuRs1.EOF
                        If Not IsNull(SubMenuRs1.Fields("description")) Then
                            lstCajas.AddItem SubMenuRs1.Fields("description")
                        End If
                        SubMenuRs1.MoveNext
                    Wend
                End If
            End If
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'USUARIOS
        '======================================================================================
        'Si es utilizado en usuarios
        If FUsuarios.Visible = True Then
            On Error GoTo err
            'eliminar usuario
            If InUusario = 1 And SubMenuRs1.Fields("user_id") <> 1 Then
                'borramos responsabilidades
                With DeleteFromTable
                    .CommandText = "DELETE FROM fnd_user_resp_groups_direct                 " & _
                                   " WHERE user_id = '" & SubMenuRs1.Fields("user_id") & "';"
                    .ActiveConnection = Cn
                    .Execute
                End With
                'borramos usuario
                With DeleteFromTable
                    .CommandText = "DELETE FROM fnd_user                                    " & _
                                   " WHERE user_id = '" & SubMenuRs1.Fields("user_id") & "';"
                    .ActiveConnection = Cn
                    .Execute
                End With
                With SubMenuRs1
                    If .State = 1 Then .Close
                        .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                Exit Sub
            End If
            'eliminar responsabilidad
            If InUusario = 2 And SubMenuRs1.Fields("user_id") <> 1 Then
                'Si no se selecciono nada en la lista
                If lstusuarios.Text = "" Then
                    MsgBox "Seleccionar alguna responsabilidad", vbOKOnly, "Informacion"
                'Eliminamos el registro
                Else
                    InInteger = GetRespId(lstusuarios.Text)
                    With DeleteFromTable
                        .CommandText = "DELETE FROM fnd_user_resp_groups_direct             " & _
                                       " WHERE user_id = " & SubMenuRs1.Fields("user_id") & _
                                       "   AND responsibility_id = " & InInteger & ";       "
                        .ActiveConnection = Cn
                        .Execute
                    End With
                    'limpiamos la lista y cargamos los datos actualizados
                    lstusuarios.Clear
                    With SubMenuRs2
                        If .State = 1 Then .Close
                            .Open SubMenuSt2, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    While Not SubMenuRs2.EOF
                        If Not IsNull(SubMenuRs2.Fields("description")) Then
                            lstusuarios.AddItem SubMenuRs2.Fields("description")
                        End If
                        SubMenuRs2.MoveNext
                    Wend
                    SubMenuRs2.Close
                End If
                'Salimos del procedimiento
                Exit Sub
            End If
        End If
        
        '======================================================================================
        'CREACION DE REPORTES
        '======================================================================================
        'Si es utilizado en creacion de reportes
        If FCreacionReportes.Visible = True Then
            On Error GoTo err
            'eliminar usuario
            'borramos las asignasiones
            With DeleteFromTable
                .CommandText = "DELETE FROM fnd_request_group_units                                         " & _
                               " WHERE request_unit_name = '" & SubMenuRs1.Fields("request_unit_name") & "';"
                .ActiveConnection = Cn
                .Execute
            End With
            'borramos el reporte
            With DeleteFromTable
                .CommandText = "DELETE FROM fnd_request_headers                                                 " & _
                                " WHERE request_header_id = '" & SubMenuRs1.Fields("request_header_id") & "';   "
                .ActiveConnection = Cn
                .Execute
            End With
            With SubMenuRs1
                If .State = 1 Then .Close
                    .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            Exit Sub
        End If
err:
        'Si hay error Cerramos el recorset y salimos del procedimiento
        With SubMenuRs2
            If .State = 1 Then .Close
        End With
        Exit Sub
End Sub
Private Sub CGuardar_Click()
    '------------------------------------------------------------------------------------------
    'variables
    '------------------------------------------------------------------------------------------
    Dim stString As String
    Dim stString2 As String
    Dim InsertIntoTable As New ADODB.Command
    Dim rsRecordset As New ADODB.Recordset
    Dim InInteger As Integer
    Dim i As Integer
    
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'GRUPOS DE CONCURRENTES
        '======================================================================================
        ' Si es usado en Grupos de concurrentes
        If FGruposConcurrentes.Visible = True Then
            'Si no seleccionamos ningun elemento de la lista
            If lstBuscar.Text = "" Then
                MsgBox "Seleccionar algun concurrente", vbOKOnly, "Informacion"
            'Insertamos el registro
            Else
                With InsertIntoTable
                    stString = Getrequest_unit_name(lstBuscar.Text)
                    .CommandText = "INSERT INTO fnd_request_group_units(        " & _
                                   "request_group_id,                           " & _
                                   "request_unit_name,                          " & _
                                   "last_update_date,                           " & _
                                   "last_updated_by,                            " & _
                                   "creation_date,                              " & _
                                   "created_by,                                 " & _
                                   "description                                 " & _
                                   ")VALUES(                                    " & _
                                    SubMenuRs1.Fields("request_group_id") & ",  " & _
                                   "'" & stString & "',                         " & _
                                   "GETDATE(),                                  " & _
                                    InUserId & ",                               " & _
                                   "GETDATE(),                                  " & _
                                    InUserId & ",                               " & _
                                   "'" & lstBuscar.Text & " '                   " & _
                                   ");                                          "
                    .ActiveConnection = Cn
                    .Execute
                End With
                'Habilitamos/deshabilitamos los botones
                CNuevo.Enabled = True
                CGuardar.Enabled = False
                CEliminar.Enabled = True
                'limpiamos los campos
                txtBuscar.Text = ""
                lstBuscar.Clear
                'Habilitamos el frame y ocultamos el frame de busqueda
                FGruposConcurrentes.Enabled = True
                FBuscar.Visible = False
                'llenamos la lista con los registros actualizados
                LGruposConcurrentes.Clear
                With SubMenuRs2
                    If .State = 1 Then .Close
                        .Open SubMenuSt2, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not SubMenuRs2.EOF
                    If Not IsNull(SubMenuRs2.Fields("description")) Then
                        LGruposConcurrentes.AddItem SubMenuRs2.Fields("description")
                    End If
                    SubMenuRs2.MoveNext
                Wend
                SubMenuRs2.Close
            End If
            Exit Sub
        End If
        
        '======================================================================================
        'CAJAS
        '======================================================================================
        ' Si es usado en cajas
        If FCajas.Visible = True Then
            'Si el textbox esta en blanco
            If txtCajas.Text = "" Then
                MsgBox "Asignar algun nombre a la caja", vbOKOnly, "Informacion"
            Else
                stString = "SELECT count(*) as existe                       " & _
                           "  FROM fnd_cajas                                " & _
                           " WHERE description = '" & txtCajas.Text & "';   "
                 With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                InInteger = rsRecordset.Fields("existe")
                rsRecordset.Close
                If InInteger <> 0 Then
                    MsgBox "Ese nombre de caja ya esta en uso", vbOKOnly, "Informacion"
                Else
                    With InsertIntoTable
                        .CommandText = "INSERT INTO fnd_cajas       " & _
                                       " (description)              " & _
                                       "Values                      " & _
                                       " ('" & txtCajas.Text & "'); "
                        .ActiveConnection = Cn
                        .Execute
                    End With
                    'limpiamos los campos
                    txtCajas.Text = ""
                    lstCajas.Clear
                    'llenamos la lista con los registros actualizados
                    With SubMenuRs1
                        If .State = 1 Then .Close
                            .Open SubMenuSt1, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    While Not SubMenuRs1.EOF
                        If Not IsNull(SubMenuRs1.Fields("description")) Then
                            lstCajas.AddItem SubMenuRs1.Fields("description")
                        End If
                        SubMenuRs1.MoveNext
                    Wend
                    SubMenuRs1.Close
                End If
            End If
            Exit Sub
        End If
        
        '======================================================================================
        'USUARIOS
        '======================================================================================
        ' Si es usado en usuarios
        If FUsuarios.Visible = True Then
            'si es usuario
            If InUusario = 1 Then
                If txtAltaUsuarios(0).Text = "" Or txtAltaUsuarios(1).Text = "" Then
                    MsgBox "Llenar el nombre de usuario y contraseña", vbOKOnly, "Informacion"
                Else
                    stString = "SELECT count(*) as existe                               " & _
                               "  FROM fnd_user                                         " & _
                               " WHERE user_name = '" & txtAltaUsuarios(0).Text & "';   "
                     With rsRecordset
                        If .State = 1 Then .Close
                            .Open stString, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    InInteger = rsRecordset.Fields("existe")
                    rsRecordset.Close
                    If InInteger <> 0 Then
                        MsgBox "Ese nombre de usuario ya esta en uso", vbOKOnly, "Informacion"
                    Else
                        InInteger = Getcaja_id(txtAltaUsuarios(2).Text)
                        With InsertIntoTable
                            .CommandText = "INSERT INTO fnd_user                " & _
                                            "(user_name,                        " & _
                                            "last_update_date,                  " & _
                                            "last_updated_by,                   " & _
                                            "creation_date,                     " & _
                                            "created_by,                        " & _
                                            "last_update_login,                 " & _
                                            "encrypted_user_password,           " & _
                                            "caja,                              " & _
                                            "start_date)                        " & _
                                           "Values                              " & _
                                            "('" & txtAltaUsuarios(0).Text & "',  " & _
                                            "GETDATE(),                         " & _
                                            InUserId & ",                       " & _
                                            "GETDATE(),                         " & _
                                            InUserId & ",                       " & _
                                            InUserId & ",                       " & _
                                            "'" & txtAltaUsuarios(1).Text & "', " & _
                                            InInteger & ",                      " & _
                                            "GETDATE());                        "
                            .ActiveConnection = Cn
                            .Execute
                        End With
                        SubMenuRs1.Requery
                        SubMenuRs1.MoveFirst
                        txtusuarios(0).Visible = True
                        txtusuarios(1).Visible = True
                        txtusuarios(2).Visible = True
                        txtAltaUsuarios(0).Visible = False
                        txtAltaUsuarios(1).Visible = False
                        txtAltaUsuarios(2).Visible = False
                        txtAltaUsuarios(0).Text = ""
                        txtAltaUsuarios(1).Text = ""
                        txtAltaUsuarios(2).Text = ""
                        txtusuarios(0).SetFocus
                        Exit Sub
                    End If
                End If
                Exit Sub
            End If
            'si es responsabilidad
            If InUusario = 2 Then
                'Si no seleccionamos ningun elemento de la lista
                If lstBuscar.Text = "" Then
                    MsgBox "Seleccionar alguna responsabilidad", vbOKOnly, "Informacion"
                'Insertamos el registro
                Else
                    With InsertIntoTable
                        InInteger = GetRespId(lstBuscar.Text)
                        .CommandText = "INSERT INTO fnd_user_resp_groups_direct(    " & _
                                       "    user_id,                                " & _
                                       "    responsibility_id,                      " & _
                                       "    start_date,                             " & _
                                       "    created_by,                             " & _
                                       "    creation_date,                          " & _
                                       "    last_updated_by,                        " & _
                                       "    last_update_date                        " & _
                                       ")VALUES(                                    " & _
                                            SubMenuRs1.Fields("user_id") & ",       " & _
                                            InInteger & ",                          " & _
                                       "    GETDATE(),                              " & _
                                            InUserId & ",                           " & _
                                       "    GETDATE(),                              " & _
                                            InUserId & ",                           " & _
                                       "    GETDATE()                               " & _
                                       ");                                          "
                        .ActiveConnection = Cn
                        .Execute
                    End With
                    'Habilitamos/deshabilitamos los botones
                    CNuevo.Enabled = True
                    CGuardar.Enabled = False
                    CEliminar.Enabled = True
                    'limpiamos los campos
                    txtBuscar.Text = ""
                    lstBuscar.Clear
                    'Habilitamos el frame y ocultamos el frame de busqueda
                    FUsuarios.Enabled = True
                    FBuscar.Visible = False
                    'llenamos la lista con los registros actualizados
                    lstusuarios.Clear
                    With SubMenuRs2
                        If .State = 1 Then .Close
                            .Open SubMenuSt2, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    While Not SubMenuRs2.EOF
                        If Not IsNull(SubMenuRs2.Fields("description")) Then
                            lstusuarios.AddItem SubMenuRs2.Fields("description")
                        End If
                        SubMenuRs2.MoveNext
                    Wend
                    SubMenuRs2.Close
                End If
                Exit Sub
            End If
        End If
        
        '======================================================================================
        'CREACION DE REPORTES
        '======================================================================================
        ' Si es usado en creacion de reportes
        If FCreacionReportes.Visible = True Then
            InsertIntoTable.CommandText = "INSERT INTO fnd_request_headers                      " & _
                                                    "(request_unit_name,                        " & _
                                                    "last_update_date,                          " & _
                                                    "last_updated_by,                           " & _
                                                    "creation_date,                             " & _
                                                    "created_by,                                " & _
                                                    "description,                               " & _
                                                    "query,                                     " & _
                                                    "parametros)                                " & _
                                                  "Values                                       " & _
                                                    "('" & txtNCreacionReportes(0).Text & "',   " & _
                                                    "GETDATE(),                                 " & _
                                                    InUserId & ",                               " & _
                                                    "GETDATE(),                                 " & _
                                                    InUserId & ",                               " & _
                                                    "'" & txtNCreacionReportes(1).Text & "',    " & _
                                                    "'" & txtNCreacionReportes(2).Text & "',    " & _
                                                    cbNCreacionReportes.Value & ");             "
            InsertIntoTable.ActiveConnection = Cn
            InsertIntoTable.Execute
            SubMenuRs1.Requery
            SubMenuRs1.MoveFirst
            txtCreacionReportes(0).Visible = True
            txtCreacionReportes(1).Visible = True
            txtCreacionReportes(2).Visible = True
            cbCreacionReportes.Visible = True
            txtNCreacionReportes(0).Visible = False
            txtNCreacionReportes(1).Visible = False
            txtNCreacionReportes(2).Visible = False
            cbNCreacionReportes.Visible = False
            txtNCreacionReportes(0).Text = ""
            txtNCreacionReportes(1).Text = ""
            txtNCreacionReportes(2).Text = ""
            cbNCreacionReportes.Value = 0
            txtCreacionReportes(0).SetFocus
            Exit Sub
        End If
        
        '======================================================================================
        'EJECUCION DE REPORTES
        '======================================================================================
        If FEjecutarReporte.Visible = True And FEjecutarReporte.Enabled = False Then
            If lstBuscar.Text <> "" Then
                txtEjecutarReporte.Text = lstBuscar.Text
                FEjecutarReporte.Enabled = True
                FBuscar.Visible = False
                CBuscar.Enabled = True
            Else
                MsgBox "Seleccionar algun reporte", vbOKOnly, "Advertencia"
            End If
            Exit Sub
        End If
        ' Si es usado en ejecucion de reportes
        If FEjecutarReporte.Visible = True Then
            If txtEjecutarReporte.Text <> "" Then
                'cuadro de dialogo para guardar el archivo
                With CDReportes
                    .DialogTitle = "Elija la ubicacion para guardar el archivo"
                    .Filter = "Archivos de excel XLS|*.xls"
                    .ShowSave
                    If .FileName = "" Then
                        'salimos ya que no se ha escrito ningún nombre de archivo
                        'o se seleccionó cancelar, por lo tanto FileName es una cadena vacía
                        Exit Sub
                    End If
                End With
                'seleccionamos el reporte
                stString = "SELECT *                                                " & _
                           "  FROM fnd_request_headers                              " & _
                           " WHERE '" & txtEjecutarReporte.Text & "' = description; "
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                InInteger = rsRecordset.Fields("parametros")
                stString2 = rsRecordset.Fields("query")
                MsgBox InInteger
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString2, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                        If InInteger = 1 Then
                            .Filter = "Fecha >= '" & Form1.DTPEjecutarReporte(0).Value & "' and Fecha <= '" & Form1.DTPEjecutarReporte(1).Value & "'"
                        End If
                End With
                'creamos el archivo
                Dim N As Long, sTemp As String
                Open CDReportes.FileName For Output As #1
                    For N = 0 To rsRecordset.Fields.Count - 1
                        sTemp = sTemp & rsRecordset.Fields(N).Name & IIf(N = rsRecordset.Fields.Count - 1, vbNullString, vbTab)
                    Next N
                    Print #1, sTemp
                    sTemp = vbNullString
                    rsRecordset.MoveFirst
                    Do Until rsRecordset.EOF
                        For N = 0 To rsRecordset.Fields.Count - 1
                            sTemp = sTemp & rsRecordset(N) & IIf(N = rsRecordset.Fields.Count - 1, vbNullString, vbTab)
                        Next N
                        Print #1, sTemp
                        sTemp = vbNullString
                        rsRecordset.MoveNext
                    Loop
                Close #1
                'abrimos el reporte
                Dim XL As New Excel.Application 'Crea el objeto excel
                XL.Workbooks.Open CDReportes.FileName, , False 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
                XL.Visible = True
                XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada.
                'limpiamos los campos
                txtEjecutarReporte.Text = ""
                For i = 0 To 1
                    DTPEjecutarReporte(i).Value = Date
                Next i
                CDReportes.FileName = ""
                txtEjecutarReporte.SetFocus
                Exit Sub
            Else
                MsgBox "Seleccione algun reporte", vbOKOnly, "Advertencia"
                Exit Sub
            End If
        End If
        
    '------------------------------------------------------------------------------------------
    'INVENTARIOS
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'ARTICULOS
        '======================================================================================
        ' Si es usado en Articulos
        If FArticulos.Visible = True Then
            'si es articulo
            If StTipoGuardarArticulo = "Articulos" Then
                If txtNArticulo(0).Text <> "" And txtNArticulo(1).Text <> "" And txtNArticulo(2).Text <> "" And txtNArticulo(3).Text <> "" And txtNArticulo(4).Text <> "" And txtNArticulo(5).Text <> "" Then
                    stString = "SELECT COUNT(*) as existe" & _
                               "  FROM mtl_system_items_b" & _
                               " WHERE segment1 ='" & txtNArticulo(0).Text & "';"
                    With rsRecordset
                        MsgBox stString
                        If .State = 1 Then .Close
                            .Open stString, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    InInteger = rsRecordset.Fields("existe")
                    rsRecordset.Close
                    If InInteger = 0 Then
                        InInteger = Getcategory_id(txtNArticulo(4).Text)
                        With InsertIntoTable
                            .CommandText = "INSERT INTO erp.dbo.mtl_system_items_b  " & _
                                           "      (description                      " & _
                                           "      ,segment1                         " & _
                                           "      ,uom                              " & _
                                           "      ,item_cost                        " & _
                                           "      ,lot_control                      " & _
                                           "      ,category_id                      " & _
                                           "      ,tax_rate                         " & _
                                           "      ,last_update_date                 " & _
                                           "      ,last_updated_by                  " & _
                                           "      ,creation_date                    " & _
                                           "      ,created_by)                      " & _
                                           "Values                                  " & _
                                           "      ('" & txtNArticulo(1).Text & "'   " & _
                                           "      ,'" & txtNArticulo(0).Text & "'   " & _
                                           "      ,'" & txtNArticulo(2).Text & "'   " & _
                                           "      ,'" & txtNArticulo(3).Text & "'   " & _
                                           "      ," & cbNArticulos.Value & _
                                           "      ," & InInteger & _
                                           "      ,'" & txtNArticulo(5).Text & "'   " & _
                                           "      ,GETDATE()                        " & _
                                           "      ," & InUserId & _
                                           "      ,GETDATE()" & _
                                           "      ," & InUserId & ");               "
                            .ActiveConnection = Cn
                            .Execute
                        End With
                        txtNArticulo(0).Text = ""
                        txtNArticulo(1).Text = ""
                        txtNArticulo(2).Text = ""
                        txtNArticulo(3).Text = ""
                        txtNArticulo(4).Text = ""
                        txtNArticulo(5).Text = ""
                        cbNArticulos.Value = 0
                        txtNArticulo(0).SetFocus
                    Else
                        MsgBox "El codigo esta siendo utilizado por otro articulo", vbOKOnly, "Advertencia"
                    End If
                Else
                    MsgBox "Debe llenar todos los campos", vbOKOnly, "Advertencia"
                End If
                Exit Sub
            End If
            'si es udm
            If StTipoGuardarArticulo = "UDM" Then
                If lstBuscar.Text <> "" Then
                    txtNArticulo(2).Text = Mid(lstBuscar.Text, 1, 2)
                    FBuscar.Visible = False
                    FArticulos.Enabled = True
                    txtBuscar.Text = ""
                    lstBuscar.Clear
                Else
                    MsgBox "Seleccionar una unidad de medida", vbOKOnly, "Advertencia"
                End If
                StTipoGuardarArticulo = "Articulos"
                Exit Sub
            End If
            'si es categoria
            If StTipoGuardarArticulo = "Categoria" Then
                If lstBuscar.Text <> "" Then
                    txtNArticulo(4).Text = lstBuscar.Text
                    FBuscar.Visible = False
                    FArticulos.Enabled = True
                    txtBuscar.Text = ""
                    lstBuscar.Clear
                Else
                    MsgBox "Seleccionar una categoria", vbOKOnly, "Advertencia"
                End If
                StTipoGuardarArticulo = "Articulos"
                Exit Sub
            End If
        End If
End Sub
Private Sub CNuevo_Click()
    '------------------------------------------------------------------------------------------
    'variables
    '------------------------------------------------------------------------------------------
    Dim stString As String
    Dim rsRecordset As New ADODB.Recordset
    Dim i As Integer
    
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'GRUPOS DE CONCURRENTES
        '======================================================================================
        If FGruposConcurrentes.Visible = True Then
            On Error GoTo err
            'habilitamos el boton guardar
            CGuardar.Enabled = True
            'deshabilitamos los botones
            CNuevo.Enabled = False
            CEliminar.Enabled = False
            'deshabilitamos el frame
            FGruposConcurrentes.Enabled = False
            'mostramos el frame de busqueda
            FBuscar.Visible = True
            'llenamos la lista con los concurrentes existentes no asignados
            stString = "SELECT frh.description as description                                                       " & _
                       "  FROM fnd_request_headers frh                                                              " & _
                       " WHERE NOT EXISTS (SELECT *                                                                 " & _
                       "                     FROM fnd_request_group_units frgu                                      " & _
                       "                    WHERE frgu.request_group_id = " & SubMenuRs1.Fields("request_group_id") & _
                       "                      AND frh.request_unit_name = frgu.request_unit_name                    " & _
                       "                  );                                                                        "
            With rsRecordset
                If .State = 1 Then .Close
                    .Open stString, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not rsRecordset.EOF
                If Not IsNull(rsRecordset.Fields("description")) Then
                    lstBuscar.AddItem rsRecordset.Fields("description")
                End If
                rsRecordset.MoveNext
            Wend
            rsRecordset.Close
            'Salimos del procedimiento
            Exit Sub
        End If
        
        '======================================================================================
        'USUARIOS
        '======================================================================================
        If FUsuarios.Visible = True Then
            On Error GoTo err1
            If InUusario = 2 Then
                'habilitamos el boton guardar
                CGuardar.Enabled = True
                'deshabilitamos los botones
                CNuevo.Enabled = False
                CEliminar.Enabled = False
                'deshabilitamos el frame
                FUsuarios.Enabled = False
                'mostramos el frame de busqueda
                FBuscar.Visible = True
                'llenamos la lista con los concurrentes existentes no asignados
                stString = "SELECT fr.description as description                                                        " & _
                           "  FROM fnd_responsibility fr                                                                " & _
                           " WHERE NOT EXISTS (SELECT *                                                                 " & _
                           "                     FROM fnd_user_resp_groups_direct furgd                                 " & _
                           "                    WHERE furgd.user_id = " & SubMenuRs1.Fields("user_id") & _
                           "                      AND fr.responsibility_id = furgd.responsibility_id                     " & _
                           "                  );                                                                        "
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not rsRecordset.EOF
                    If Not IsNull(rsRecordset.Fields("description")) Then
                        lstBuscar.AddItem rsRecordset.Fields("description")
                    End If
                    rsRecordset.MoveNext
                Wend
                rsRecordset.Close
            End If
            'Salimos del procedimiento
            Exit Sub
        End If
        
    '------------------------------------------------------------------------------------------
    'INVENTARIOS
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'ARTICULOS
        '======================================================================================
        If FArticulos.Visible = True Then
            On Error Resume Next
            txtArticulo(0).Visible = False
            txtArticulo(1).Visible = False
            txtArticulo(2).Visible = False
            txtArticulo(3).Visible = False
            txtArticulo(4).Visible = False
            txtArticulo(5).Visible = False
            txtNArticulo(0).Visible = True
            txtNArticulo(1).Visible = True
            txtNArticulo(2).Visible = True
            txtNArticulo(3).Visible = True
            txtNArticulo(4).Visible = True
            txtNArticulo(5).Visible = True
            CbtArticulos(0).Visible = True
            CbtArticulos(1).Visible = True
            cbArticulos.Visible = False
            cbNArticulos.Visible = True
            CNuevo.Enabled = False
            CBuscar.Enabled = False
            txtNArticulo(0).SetFocus
            Exit Sub
        End If
err:
        'si hay error cerramos el recordset y salimos del procedimiento
        With rsRecordset
            If .State = 1 Then .Close
        End With
        'si la lista no contiene nada
        MsgBox "No hay concurrentes disponibles", vbOKOnly, "Atencion"
        'habilitamos/deshabilitamos los botones
        CNuevo.Enabled = True
        CGuardar.Enabled = False
        CEliminar.Enabled = True
        'cerramos el cuadro de busqueda y habilitamos el frame
        FBuscar.Visible = False
        FGruposConcurrentes.Enabled = True
        'salimos del procedimiento
        Exit Sub
err1:
        'si hay error cerramos el recordset y salimos del procedimiento
        With rsRecordset
            If .State = 1 Then .Close
        End With
        'si la lista no contiene nada
        MsgBox "No hay responsabilidades disponibles", vbOKOnly, "Atencion"
        'habilitamos/deshabilitamos los botones
        CNuevo.Enabled = True
        CGuardar.Enabled = False
        CEliminar.Enabled = True
        'cerramos el cuadro de busqueda y habilitamos el frame
        FBuscar.Visible = False
        FUsuarios.Enabled = True
        'salimos del procedimiento
        Exit Sub
End Sub
Private Sub CBuscar_Click()
    '------------------------------------------------------------------------------------------
    'variables
    '------------------------------------------------------------------------------------------
    Dim stString As String
    Dim rsRecordset As New ADODB.Recordset
    Dim i As Integer
    
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'EJECUTAR REPORTES
        '======================================================================================
        If FEjecutarReporte.Visible = True Then
            On Error GoTo err
            'habilitamos el boton guardar
            CBuscar.Enabled = False
            CGuardar.Enabled = True
            'deshabilitamos el frame
            FEjecutarReporte.Enabled = False
            'mostramos el frame de busqueda
            FBuscar.Visible = True
            txtBuscar.Text = ""
            lstBuscar.Clear
            'llenamos la lista con los concurrentes existentes asignados
            stString = "SELECT frh.description                                  " & _
                       "  FROM fnd_request_headers     frh,                     " & _
                       "       fnd_request_group_units frgu,                    " & _
                       "       fnd_request_groups      frg,                     " & _
                       "       fnd_responsibility fr                            " & _
                       " WHERE frh.request_unit_name = frgu.request_unit_name   " & _
                       "   AND frgu.request_group_id = frg.request_group_id     " & _
                       "   AND frg.request_group_id  = fr.request_group_id      " & _
                       "   AND fr.responsibility_id  = " & InRespId & ";        "
            With rsRecordset
                If .State = 1 Then .Close
                    .Open stString, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not rsRecordset.EOF
                If Not IsNull(rsRecordset.Fields("description")) Then
                    lstBuscar.AddItem rsRecordset.Fields("description")
                End If
                rsRecordset.MoveNext
            Wend
            rsRecordset.Close
            'Salimos del procedimiento
            Exit Sub
        End If
        
    '------------------------------------------------------------------------------------------
    'INVENTARIOS
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'ARTICULOS
        '======================================================================================
        If FArticulos.Visible = True Then
            On Error GoTo err1
            'habilitamos el boton guardar
            CBuscar.Enabled = False
            CNuevo.Enabled = False
            'deshabilitamos el frame
            FArticulos.Enabled = False
            'mostramos el frame de busqueda
            FBuscar.Visible = True
            txtBuscar.Text = ""
            lstBuscar.Clear
            'llenamos la lista
            StTipoGuardarArticulo = "Articulos"
            stString = "SELECT inventory_item_id,                                                               " & _
                       "       right(SUBSTRING(segment1,1,10),10)+'          '+' - '+description as description " & _
                       "  FROM mtl_system_items_b;                                                              "
            With rsRecordset
                If .State = 1 Then .Close
                    .Open stString, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not rsRecordset.EOF
                If Not IsNull(rsRecordset.Fields("description")) Then
                    lstBuscar.AddItem rsRecordset.Fields("description")
                End If
                rsRecordset.MoveNext
            Wend
            rsRecordset.Close
            'Salimos del procedimiento
            Exit Sub
        End If
err:
        'si hay error cerramos el recordset y salimos del procedimiento
        With rsRecordset
            If .State = 1 Then .Close
        End With
        'si la lista no contiene nada
        MsgBox "No hay concurrentes disponibles", vbOKOnly, "Atencion"
        'habilitamos/deshabilitamos los botones
        CGuardar.Enabled = False
        CBuscar.Enabled = True
        'cerramos el cuadro de busqueda y habilitamos el frame
        FBuscar.Visible = False
        FEjecutarReporte.Enabled = True
        'salimos del procedimiento
        Exit Sub
err1:
        'si hay error cerramos el recordset y salimos del procedimiento
        With rsRecordset
            If .State = 1 Then .Close
        End With
        'si la lista no contiene nada
        MsgBox "No hay opciones disponibles", vbOKOnly, "Atencion"
        'habilitamos/deshabilitamos los botones
        CBuscar.Enabled = True
        CNuevo.Enabled = True
        'cerramos el cuadro de busqueda y habilitamos el frame
        FBuscar.Visible = False
        FArticulos.Enabled = True
        'salimos del procedimiento
        Exit Sub
End Sub




'==============================================================================================

'B  U   S   C   A   D   O   R

'==============================================================================================
Private Sub txtBuscar_Change()
    'buscador
    'variables
    Dim stString As String
    Dim rsRecordset As New ADODB.Recordset
    'limpiamos la lista
    lstBuscar.Clear
    '------------------------------------------------------------------------------------------
    'ADMINISTRADOR DEL SISTEMA
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'GRUPOS DE CONCURRENTES
        '======================================================================================
        'si es usado por Grupos de concurrentes
        If FGruposConcurrentes.Visible = True Then
            On Error GoTo err
            'si no esta en blanco lo filtramos
            If txtBuscar <> "" Then
                stString = "SELECT fr.description as description                                        " & _
                           "  FROM fnd_responsibility fr,                                               " & _
                           " WHERE NOT EXISTS (SELECT *                                                 " & _
                           "                     FROM fnd_user_resp_groups_direct furgd                 " & _
                           "                    WHERE furgd.user_id = " & SubMenuRs1.Fields("user_id") & _
                           "                      AND fr.responsibility_id = frgu.responsibility_id)    " & _
                           "   AND fr.description like '%" & txtBuscar & "%';                           "
            Else
                'si no mostramos todos los resultados
                stString = "SELECT fr.description as description                                        " & _
                           "  FROM fnd_responsibility fr,                                               " & _
                           " WHERE NOT EXISTS (SELECT *                                                 " & _
                           "                     FROM fnd_user_resp_groups_direct furgd                 " & _
                           "                    WHERE furgd.user_id = " & SubMenuRs1.Fields("user_id") & _
                           "                      AND fr.responsibility_id = frgu.responsibility_id);   "
            End If
            'llenamos la lista
            With rsRecordset
                If .State = 1 Then .Close
                    .Open stString, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not rsRecordset.EOF
                If Not IsNull(rsRecordset.Fields("description")) Then
                    lstBuscar.AddItem rsRecordset.Fields("description")
                End If
                rsRecordset.MoveNext
            Wend
            rsRecordset.Close
            'salimos del procedimiento
            Exit Sub
            'si hay error cerramos el recordset y salimos del procedimiento
        End If
        '======================================================================================
        'USUARIOS
        '======================================================================================
        'si es usao por Grupos de concurrentes
        If FUsuarios.Visible = True Then
            On Error GoTo err
            'si no esta en blanco lo filtramos
            If txtBuscar <> "" Then
                stString = "SELECT fr.description as description                                        " & _
                           "  FROM fnd_responsibility fr,                                               " & _
                           " WHERE NOT EXISTS (SELECT *                                                 " & _
                           "                     FROM fnd_user_resp_groups_direct furgd                 " & _
                           "                    WHERE furgd.user_id = " & SubMenuRs1.Fields("user_id") & _
                           "                      AND fr.responsibility_id = frgu.responsibility_id)    " & _
                           "   AND fr.description like '%" & txtBuscar & "%';                           "
            Else
                'si no mostramos todos los resultados
                stString = "SELECT fr.description as description                                        " & _
                           "  FROM fnd_responsibility fr,                                               " & _
                           " WHERE NOT EXISTS (SELECT *                                                 " & _
                           "                     FROM fnd_user_resp_groups_direct furgd                 " & _
                           "                    WHERE furgd.user_id = " & SubMenuRs1.Fields("user_id") & _
                           "                      AND fr.responsibility_id = frgu.responsibility_id);   "
            End If
            'llenamos la lista
            With rsRecordset
                If .State = 1 Then .Close
                    .Open stString, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not rsRecordset.EOF
                If Not IsNull(rsRecordset.Fields("description")) Then
                    lstBuscar.AddItem rsRecordset.Fields("description")
                End If
                rsRecordset.MoveNext
            Wend
            rsRecordset.Close
            'salimos del procedimiento
            Exit Sub
            'si hay error cerramos el recordset y salimos del procedimiento
        End If
        
        '======================================================================================
        'EJECUCION DE REPORTES
        '======================================================================================
        'si es usado por ejecucion de reportes
        If FEjecutarReporte.Visible = True Then
            On Error GoTo err
            'si no esta en blanco lo filtramos
            If txtBuscar <> "" Then
                stString = "SELECT frh.description                                  " & _
                           "  FROM fnd_request_headers     frh,                     " & _
                           "       fnd_request_group_units frgu,                    " & _
                           "       fnd_request_groups      frg,                     " & _
                           "       fnd_responsibility fr                            " & _
                           " WHERE frh.request_unit_name = frgu.request_unit_name   " & _
                           "   AND frgu.request_group_id = frg.request_group_id     " & _
                           "   AND frg.request_group_id  = fr.request_group_id      " & _
                           "   AND fr.responsibility_id  = " & InRespId & _
                           "   AND frh.description like '%" & txtBuscar & "%';      "
            Else
                'si no mostramos todos los resultados
                stString = "SELECT frh.description                                  " & _
                           "  FROM fnd_request_headers     frh,                     " & _
                           "       fnd_request_group_units frgu,                    " & _
                           "       fnd_request_groups      frg,                     " & _
                           "       fnd_responsibility fr                            " & _
                           " WHERE frh.request_unit_name = frgu.request_unit_name   " & _
                           "   AND frgu.request_group_id = frg.request_group_id     " & _
                           "   AND frg.request_group_id  = fr.request_group_id      " & _
                           "   AND fr.responsibility_id  = " & InRespId & ";        "
            End If
            'llenamos la lista
            With rsRecordset
                If .State = 1 Then .Close
                    .Open stString, Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .MoveFirst
            End With
            While Not rsRecordset.EOF
                If Not IsNull(rsRecordset.Fields("description")) Then
                    lstBuscar.AddItem rsRecordset.Fields("description")
                End If
                rsRecordset.MoveNext
            Wend
            rsRecordset.Close
            'salimos del procedimiento
            Exit Sub
            'si hay error cerramos el recordset y salimos del procedimiento
        End If
    '------------------------------------------------------------------------------------------
    'INVENTARIOS
    '------------------------------------------------------------------------------------------
        '======================================================================================
        'ARTICULOS
        '======================================================================================
        'si es usado por articulos
        If FArticulos.Visible = True Then
            On Error GoTo err
            'si es usado en articulos
            If StTipoGuardarArticulo = "Articulos" Then
                'si no esta en blanco lo filtramos
                If txtBuscar <> "" Then
                    stString = "SELECT inventory_item_id,                                                               " & _
                               "       right(SUBSTRING(segment1,1,10),10)+'          '+' - '+description as description " & _
                               "  FROM mtl_system_items_b                                                               " & _
                               " WHERE segment1 +' - '+description like '%" & txtBuscar & "%';                          "
                Else
                    'si no mostramos todos los resultados
                    stString = "SELECT inventory_item_id,                                                               " & _
                               "       right(SUBSTRING(segment1,1,10),10)+'          '+' - '+description as description " & _
                               "  FROM mtl_system_items_b                                                               "
                End If
                'llenamos la lista
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not rsRecordset.EOF
                    If Not IsNull(rsRecordset.Fields("description")) Then
                        lstBuscar.AddItem rsRecordset.Fields("description")
                    End If
                    rsRecordset.MoveNext
                Wend
                rsRecordset.Close
                'salimos del procedimiento
                Exit Sub
            End If
            'si es usado en udm
            If StTipoGuardarArticulo = "UDM" Then
                'si no esta en blanco lo filtramos
                If txtBuscar <> "" Then
                    stString = "SELECT uom_code+' - '+description as description                " & _
                               "  FROM mtl_units_of_measure                                     " & _
                               " WHERE uom_code+' - '+description like '%" & txtBuscar & "%';   "
                Else
                    'si no mostramos todos los resultados
                    stString = "SELECT uom_code+' - '+description as description    " & _
                               "  FROM mtl_units_of_measure;                        "
                End If
                'llenamos la lista
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not rsRecordset.EOF
                    If Not IsNull(rsRecordset.Fields("description")) Then
                        lstBuscar.AddItem rsRecordset.Fields("description")
                    End If
                    rsRecordset.MoveNext
                Wend
                rsRecordset.Close
                'salimos del procedimiento
                Exit Sub
            End If
            'si es usado en categorias
            If StTipoGuardarArticulo = "Categorias" Then
                'si no esta en blanco lo filtramos
                If txtBuscar <> "" Then
                    stString = "SELECT category_id,                                 " & _
                               "       description                                  " & _
                               "  FROM mtl_item_categories                          " & _
                               " WHERE description like '" & txtBuscar.Text & "';   "
                Else
                    'si no mostramos todos los resultados
                    stString = "SELECT category_id,         " & _
                               "       description          " & _
                               "  FROM mtl_item_categories; "
                End If
                'llenamos la lista
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not rsRecordset.EOF
                    If Not IsNull(rsRecordset.Fields("description")) Then
                        lstBuscar.AddItem rsRecordset.Fields("description")
                    End If
                    rsRecordset.MoveNext
                Wend
                rsRecordset.Close
                'salimos del procedimiento
                Exit Sub
            End If
        End If
'si hay error cerramos el recordset y salimos del procedimiento
err:
        With rsRecordset
            If .State = 1 Then .Close
        End With
        Exit Sub
End Sub



'==============================================================================================

'A  D   M   I   N   I   S   T   R   A   D   O   R       D   E       S   I   S   T   E   M   A

'==============================================================================================

'----------------------------------------------------------------------------------------------
'GRUPOS DE CONCURRENTES
'----------------------------------------------------------------------------------------------
Private Sub TGruposConcurrentes_Change()
    On Error GoTo err
    'si no esta en blanco habilitamos los botones
    If FGruposConcurrentes.Visible = True Then
        If TGruposConcurrentes.Text <> "" Then
            CNuevo.Enabled = True
            CEliminar.Enabled = True
        Else
            CNuevo.Enabled = False
            CEliminar.Enabled = False
        End If
    End If
    'llenamos la lista
    LGruposConcurrentes.Clear
    SubMenuSt2 = "SELECT description                                                        " & _
                 "  FROM fnd_request_group_units                                            " & _
                 " WHERE request_group_id = " & SubMenuRs1.Fields("request_group_id") & ";  "
    With SubMenuRs2
        If .State = 1 Then .Close
            .Open SubMenuSt2, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    While Not SubMenuRs2.EOF
        If Not IsNull(SubMenuRs2.Fields("description")) Then
            LGruposConcurrentes.AddItem SubMenuRs2.Fields("description")
        End If
        SubMenuRs2.MoveNext
    Wend
    SubMenuRs2.Close
    'salimos del procedimiento
    Exit Sub
err:
    'si hay error cerramos el recordset
    With SubMenuRs2
        If .State = 1 Then .Close
    End With
End Sub
Private Sub TGruposConcurrentes_KeyDown(KeyCode As Integer, Shift As Integer)
    'nos movemos entre registros con la flecha abajo o arriba
    On Error GoTo err
    If KeyCode = 40 Then
        SubMenuRs1.MoveNext
        Exit Sub
    End If
    If KeyCode = 38 Then
        SubMenuRs1.MovePrevious
        Exit Sub
    End If
'si hay error vamos al primer registro
err:
    SubMenuRs1.MoveFirst
End Sub
Private Sub txtEjecutarReporte_Change()
    If txtEjecutarReporte.Text = "" And FEjecutarReporte.Visible = True Then
        CGuardar.Enabled = False
    Else
        CGuardar.Enabled = True
    End If
End Sub




'----------------------------------------------------------------------------------------------
'USUARIOS
'----------------------------------------------------------------------------------------------
Private Sub txtusuarios_Change(Index As Integer)
    On Error GoTo err
    Select Case Index
        Case 0
            'si no esta en blanco habilitamos los botones
            If FUsuarios.Visible = True Then
                If txtusuarios(0).Text <> "" Then
                    CNuevo.Enabled = True
                    CEliminar.Enabled = True
                    CGuardar.Enabled = False
                    txtusuarios(0).Visible = True
                    txtusuarios(1).Visible = True
                    txtusuarios(2).Visible = True
                    txtAltaUsuarios(0).Visible = False
                    txtAltaUsuarios(1).Visible = False
                    txtAltaUsuarios(2).Visible = False
                    'llenamos la lista
                    lstusuarios.Clear
                    SubMenuSt2 = "SELECT t1.description                                     " & _
                                 "  FROM  fnd_responsibility          t1,                   " & _
                                 "        fnd_user_resp_groups_direct t2                    " & _
                                 " WHERE t1.responsibility_id = t2.responsibility_id        " & _
                                 "   AND t2.user_id = " & SubMenuRs1.Fields("user_id") & "; "
                    With SubMenuRs2
                        If .State = 1 Then .Close
                            .Open SubMenuSt2, Cn, adOpenStatic, adLockOptimistic
                            .Requery
                            .MoveFirst
                    End With
                    While Not SubMenuRs2.EOF
                        If Not IsNull(SubMenuRs2.Fields("description")) Then
                            lstusuarios.AddItem SubMenuRs2.Fields("description")
                        End If
                        SubMenuRs2.MoveNext
                    Wend
                    SubMenuRs2.Close
                    txtusuarios(0).SetFocus
                    'salimos del procedimiento
                    Exit Sub
                Else
                    lstusuarios.Clear
                    CGuardar.Enabled = True
                    CNuevo.Enabled = False
                    CEliminar.Enabled = False
                    txtusuarios(0).Visible = False
                    txtusuarios(1).Visible = False
                    txtusuarios(2).Visible = False
                    txtAltaUsuarios(0).Visible = True
                    txtAltaUsuarios(1).Visible = True
                    txtAltaUsuarios(2).Visible = True
                    txtAltaUsuarios(0).SetFocus
                    Exit Sub
                End If
            End If
    End Select
err:
    'si hay error cerramos el recordset
    With SubMenuRs2
        If .State = 1 Then .Close
    End With
End Sub
Private Sub txtusuarios_Click(Index As Integer)
    InUusario = 1
End Sub
Private Sub lstusuarios_Click()
    InUusario = 2
End Sub
Private Sub lstusuarios_GotFocus()
    InUusario = 2
End Sub
Private Sub txtAltaUsuarios_Change(Index As Integer)
    InUusario = 1
End Sub
Private Sub txtusuarios_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'nos movemos entre registros con la flecha abajo o arriba
    On Error GoTo err
    Select Case Index
        Case 0
            If KeyCode = 40 Then
                SubMenuRs1.MoveNext
                Exit Sub
            End If
            If KeyCode = 38 Then
                SubMenuRs1.MovePrevious
                Exit Sub
            End If
    End Select
'si hay error vamos al primer registro
err:
    SubMenuRs1.MoveFirst
    CGuardar.Enabled = False
End Sub

'----------------------------------------------------------------------------------------------
'CREACION DE REPORTES
'----------------------------------------------------------------------------------------------
Private Sub txtCreacionReportes_Change(Index As Integer)
    Select Case Index
        Case 0
            'si no esta en blanco habilitamos los botones
            If FCreacionReportes.Visible = True Then
                If txtCreacionReportes(0).Text <> "" Then
                    CEliminar.Enabled = True
                    CGuardar.Enabled = False
                    txtCreacionReportes(0).Visible = True
                    txtCreacionReportes(1).Visible = True
                    txtCreacionReportes(2).Visible = True
                    cbCreacionReportes.Visible = True
                    txtNCreacionReportes(0).Visible = False
                    txtNCreacionReportes(1).Visible = False
                    txtNCreacionReportes(2).Visible = False
                    cbNCreacionReportes.Visible = False
                    Exit Sub
                Else
                    CGuardar.Enabled = True
                    CEliminar.Enabled = False
                    txtCreacionReportes(0).Visible = False
                    txtCreacionReportes(1).Visible = False
                    txtCreacionReportes(2).Visible = False
                    cbCreacionReportes.Visible = False
                    txtNCreacionReportes(0).Visible = True
                    txtNCreacionReportes(1).Visible = True
                    txtNCreacionReportes(2).Visible = True
                    cbNCreacionReportes.Visible = True
                    txtNCreacionReportes(0).SetFocus
                    Exit Sub
                End If
            End If
    End Select
End Sub
Private Sub txtCreacionReportes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'nos movemos entre registros con la flecha abajo o arriba
    On Error GoTo err
    Select Case Index
        Case 0
            If KeyCode = 40 Then
                SubMenuRs1.MoveNext
                Exit Sub
            End If
            If KeyCode = 38 Then
                SubMenuRs1.MovePrevious
                Exit Sub
            End If
    End Select
'si hay error vamos al primer registro
err:
    SubMenuRs1.MoveFirst
    CGuardar.Enabled = False
End Sub




'==============================================================================================

'I  N   V   E   N   T   A   R   I   O   S

'==============================================================================================

'----------------------------------------------------------------------------------------------
'ARTICULOS
'----------------------------------------------------------------------------------------------
Private Sub txtArticulo_Change(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0
            'si no esta en blanco habilitamos los botones
            If FArticulos.Visible = True Then
                If txtArticulo(0).Text <> "" Then
                    For i = 0 To 5
                        txtArticulo(i).Visible = True
                        txtNArticulo(i).Visible = False
                    Next i
                    For i = 0 To 1
                        CbtArticulos(i).Visible = False
                    Next i
                    cbArticulos.Visible = True
                    cbNArticulos.Visible = False
                    CNuevo.Enabled = True
                    Exit Sub
                Else
                    For i = 0 To 5
                        txtArticulo(i).Visible = False
                        txtNArticulo(i).Visible = True
                    Next i
                    For i = 0 To 1
                        CbtArticulos(i).Visible = True
                    Next i
                    cbArticulos.Visible = False
                    cbNArticulos.Visible = True
                    CNuevo.Enabled = False
                    txtNArticulo(0).SetFocus
                    Exit Sub
                End If
            End If
    End Select
End Sub
Private Sub txtArticulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'nos movemos entre registros con la flecha abajo o arriba
    On Error GoTo err
    Select Case Index
        Case 0
            If KeyCode = 40 Then
                SubMenuRs1.MoveNext
                Exit Sub
            End If
            If KeyCode = 38 Then
                SubMenuRs1.MovePrevious
                Exit Sub
            End If
    End Select
'si hay error vamos al primer registro
err:
    On Error Resume Next
    SubMenuRs1.MoveFirst
End Sub
Private Sub CbtArticulos_Click(Index As Integer)
    Dim stString As String
    Dim rsRecordset As New ADODB.Recordset
    On Error GoTo err
    Select Case Index
        Case 0
            StTipoBuscadorArticulo = "UDM"
            StTipoGuardarArticulo = "UDM"
            If FArticulos.Visible = True Then
                'habilitamos el boton guardar
                CBuscar.Enabled = False
                CNuevo.Enabled = False
                'deshabilitamos el frame
                FArticulos.Enabled = False
                'mostramos el frame de busqueda
                FBuscar.Visible = True
                txtBuscar.Text = ""
                lstBuscar.Clear
                'llenamos la lista
                stString = "SELECT uom_code+' - '+description as description    " & _
                           "  FROM mtl_units_of_measure;                        "
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not rsRecordset.EOF
                    If Not IsNull(rsRecordset.Fields("description")) Then
                        lstBuscar.AddItem rsRecordset.Fields("description")
                    End If
                    rsRecordset.MoveNext
                Wend
                rsRecordset.Close
                'Salimos del procedimiento
                Exit Sub
            End If
        Case 1
            StTipoBuscadorArticulo = "Categoria"
            StTipoGuardarArticulo = "Categoria"
            If FArticulos.Visible = True Then
                'habilitamos el boton guardar
                CBuscar.Enabled = False
                CNuevo.Enabled = False
                'deshabilitamos el frame
                FArticulos.Enabled = False
                'mostramos el frame de busqueda
                FBuscar.Visible = True
                txtBuscar.Text = ""
                lstBuscar.Clear
                'llenamos la lista
                stString = "SELECT category_id,         " & _
                           "       description          " & _
                           "  FROM mtl_item_categories; "
                With rsRecordset
                    If .State = 1 Then .Close
                        .Open stString, Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .MoveFirst
                End With
                While Not rsRecordset.EOF
                    If Not IsNull(rsRecordset.Fields("description")) Then
                        lstBuscar.AddItem rsRecordset.Fields("description")
                    End If
                    rsRecordset.MoveNext
                Wend
                rsRecordset.Close
                'Salimos del procedimiento
                Exit Sub
            End If
    End Select
err:
        'si hay error cerramos el recordset y salimos del procedimiento
        With rsRecordset
            If .State = 1 Then .Close
        End With
        'si la lista no contiene nada
        MsgBox "No hay opciones disponibles", vbOKOnly, "Atencion"
        'habilitamos/deshabilitamos los botones
        CBuscar.Enabled = True
        CNuevo.Enabled = True
        'cerramos el cuadro de busqueda y habilitamos el frame
        FBuscar.Visible = False
        FArticulos.Enabled = True
        'salimos del procedimiento
        Exit Sub
End Sub
