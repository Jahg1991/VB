VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Compras"
   ClientHeight    =   8805
   ClientLeft      =   2505
   ClientTop       =   2550
   ClientWidth     =   10830
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8535
      Left            =   120
      TabIndex        =   96
      Top             =   120
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox txtPeso 
         Height          =   375
         Left            =   120
         TabIndex        =   104
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   7080
         Width           =   2500
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   1
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   7080
         Width           =   2500
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   6720
         Width           =   2500
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   6360
         Width           =   2500
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   6000
         Width           =   2500
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   5640
         Width           =   2500
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   5280
         Width           =   2500
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   4920
         Width           =   2500
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   4560
         Width           =   2500
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   4200
         Width           =   2500
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   3840
         Width           =   2500
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3480
         Width           =   2500
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3120
         Width           =   2500
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2760
         Width           =   2500
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   2500
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2040
         Width           =   2500
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   6720
         Width           =   330
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   6360
         Width           =   330
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   6000
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   5640
         Width           =   330
      End
      Begin VB.CommandButton Command11 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5280
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4920
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4560
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   4200
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3840
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3480
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3120
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2760
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2400
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   330
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   6720
         Width           =   2500
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   6360
         Width           =   2500
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   6000
         Width           =   2500
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   5640
         Width           =   2500
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   5280
         Width           =   2500
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   4920
         Width           =   2500
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   4560
         Width           =   2500
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   4200
         Width           =   2500
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3840
         Width           =   2500
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3480
         Width           =   2500
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3120
         Width           =   2500
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2760
         Width           =   2500
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   2500
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2040
         Width           =   2500
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   88
         Top             =   6720
         Width           =   2000
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   82
         Top             =   6360
         Width           =   2000
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   76
         Top             =   6000
         Width           =   2000
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   70
         Top             =   5640
         Width           =   2000
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   64
         Top             =   5280
         Width           =   2000
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   58
         Top             =   4920
         Width           =   2000
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   52
         Top             =   4560
         Width           =   2000
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   46
         Top             =   4200
         Width           =   2000
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   40
         Top             =   3840
         Width           =   2000
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   34
         Top             =   3480
         Width           =   2000
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   28
         Top             =   3120
         Width           =   2000
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   22
         Top             =   2760
         Width           =   2000
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   16
         Top             =   2400
         Width           =   2000
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   2040
         Width           =   2000
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   6720
         Width           =   330
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   6360
         Width           =   330
      End
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   6000
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   5640
         Width           =   330
      End
      Begin VB.CommandButton Command11 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   5280
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   4920
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4560
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4200
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3840
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3480
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3120
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2760
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   330
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   6720
         Width           =   2500
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   6360
         Width           =   2500
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   6000
         Width           =   2500
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   5640
         Width           =   2500
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   5280
         Width           =   2500
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   4920
         Width           =   2500
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   4560
         Width           =   2500
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   4200
         Width           =   2500
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3840
         Width           =   2500
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3480
         Width           =   2500
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3120
         Width           =   2500
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2760
         Width           =   2500
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2400
         Width           =   2500
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   2500
      End
      Begin VB.CommandButton Command16 
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
         Left            =   7920
         TabIndex        =   95
         Top             =   7800
         Width           =   2535
      End
      Begin VB.TextBox Text17 
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
         Height          =   495
         Left            =   1440
         TabIndex        =   94
         Top             =   7800
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   2500
      End
      Begin VB.TextBox Text0 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   9825
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   2500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   4
         Top             =   1680
         Width           =   2000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Index           =   1
         Left            =   2630
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   2500
      End
      Begin VB.CommandButton Command0 
         Caption         =   "..."
         Height          =   495
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total:"
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
         Index           =   6
         Left            =   3720
         TabIndex        =   103
         Top             =   7080
         Width           =   1215
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
         TabIndex        =   102
         Top             =   7920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subtotal"
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
         Left            =   7920
         TabIndex        =   101
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   5040
         TabIndex        =   100
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   3000
         TabIndex        =   99
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         TabIndex        =   98
         Top             =   1320
         Width           =   2415
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
         Left            =   0
         TabIndex        =   97
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3840
      Top             =   6840
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
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   0
      Picture         =   "Form1.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
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
    Form1.Enabled = False
    
End Sub

Private Sub Command0_Click()
    
    MBProveedor
        
End Sub

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 1
            Text1(1).Text = ""
            IdArticulo(1) = 0
            Command16.Enabled = False
            MBArticulo
            
        Case 2
            
            LeerPuertoBascula
            Text1(3).Enabled = True
            Text1(3).Text = txtPeso.Text
            Command2(1).Enabled = True
            Command2(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command2_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 2
            Text2(1).Text = ""
            IdArticulo(2) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text2(3).Enabled = True
            Text2(3).Text = txtPeso.Text
            Command3(1).Enabled = True
            Command3(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command3_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 3
            Text3(1).Text = ""
            IdArticulo(3) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text3(3).Enabled = True
            Text3(3).Text = txtPeso.Text
            Command4(1).Enabled = True
            Command4(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command4_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 4
            Text4(1).Text = ""
            IdArticulo(4) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text4(3).Enabled = True
            Text4(3).Text = txtPeso.Text
            Command5(1).Enabled = True
            Command5(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command5_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 5
            Text5(1).Text = ""
            IdArticulo(5) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text5(3).Enabled = True
            Text5(3).Text = txtPeso.Text
            Command6(1).Enabled = True
            Command6(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command6_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 6
            Text6(1).Text = ""
            IdArticulo(6) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text6(3).Enabled = True
            Text6(3).Text = txtPeso.Text
            Command7(1).Enabled = True
            Command7(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command7_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 7
            Text7(1).Text = ""
            IdArticulo(7) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text7(3).Enabled = True
            Text7(3).Text = txtPeso.Text
            Command8(1).Enabled = True
            Command8(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command8_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 8
            Text8(1).Text = ""
            IdArticulo(8) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text8(3).Enabled = True
            Text8(3).Text = txtPeso.Text
            Command9(1).Enabled = True
            Command9(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command9_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 9
            Text9(1).Text = ""
            IdArticulo(9) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text9(3).Enabled = True
            Text9(3).Text = txtPeso.Text
            Command10(1).Enabled = True
            Command10(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command10_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 10
            Text10(1).Text = ""
            IdArticulo(10) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text10(3).Enabled = True
            Text10(3).Text = txtPeso.Text
            Command11(1).Enabled = True
            Command11(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command11_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 11
            Text11(1).Text = ""
            IdArticulo(11) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text11(3).Enabled = True
            Text11(3).Text = txtPeso.Text
            Command12(1).Enabled = True
            Command12(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command12_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 12
            Text12(1).Text = ""
            IdArticulo(12) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text12(3).Enabled = True
            Text12(3).Text = txtPeso.Text
            Command13(1).Enabled = True
            Command13(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command13_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 13
            Text13(1).Text = ""
            IdArticulo(13) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text13(3).Enabled = True
            Text13(3).Text = txtPeso.Text
            Command14(1).Enabled = True
            Command14(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command14_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 14
            Text14(1).Text = ""
            IdArticulo(14) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text14(3).Enabled = True
            Text14(3).Text = txtPeso.Text
            Command15(1).Enabled = True
            Command15(1).SetFocus
            
    End Select
  
End Sub

Private Sub Command15_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            LArticulo = 15
            Text15(1).Text = ""
            IdArticulo(15) = 0
            MBArticulo
            
        Case 2
            LeerPuertoBascula
            Text15(3).Enabled = True
            Text15(3).Text = txtPeso.Text
            Text17.SetFocus
            
    End Select
  
End Sub

Private Sub Command16_Click()
    
    If Text0 <> "" And Text1(1) <> "" And Text1(2) <> "" And Text1(3) <> "" And Text1(4) <> "" Then
        ICompra
    End If

End Sub

Private Sub Historialcompras_Click()
    
    OCompra
    Form3.Show
    Form1.Enabled = False
    
End Sub

Private Sub NuevaCompra_Click()
    
    NCompra
    
End Sub

Private Sub Proveedores_Click()
    
    OCompra
    CProveedores
    Form1.Enabled = False

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    RS1.Close
    RS2.Close
    RS3.Close
    RS4.Close
    RS5.Close
    
    CN.Close
    
    'aa = Shell("shutdown -s -t 00")
    
    Unload Form1
    
End Sub

Private Sub Text0_Change()
    
    If Text0 <> "" Then
        Command1(1).Enabled = True
    Else
        Command1(1).Enabled = False
    End If
    
End Sub

Private Sub Text1_Change(Index As Integer)

    On Error Resume Next

    Select Case Index
    
        Case 1
            If Text1(1).Text <> "" Then
                Text1(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text1(i).Text = ""
                    Text1(i).Enabled = False
                Next
                
                Command1(2).Enabled = False
                
                TSubtotal(1) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(1) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If

        Case 2
            If Text1(1) <> "" And Text1(2) <> "" And Text1(3) <> "" And Text1(3) <> "ERROR" Then
                Text1(4) = Round(Replace(Text1(3), "kg", "") * Text1(2), 2)
            End If
            
        Case 3
            If Text1(1) <> "" And Text1(2) <> "" And Text1(3) <> "" And Text1(3) <> "ERROR" Then
                Text1(4) = Round(Replace(Text1(3), "kg", "") * Text1(2), 2)
                TPeso(1) = Replace(Text1(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(1) = 0
            End If
            
        Case 4
            TSubtotal(1) = Text1(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text2_Change(Index As Integer)

    On Error Resume Next

    Select Case Index
    
        Case 1
            If Text2(1).Text <> "" Then
                Text2(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text2(i).Text = ""
                    Text2(i).Enabled = False
                Next
                
                Command2(2).Enabled = False
                
                TSubtotal(2) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(2) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If

        Case 2
            If Text2(1) <> "" And Text2(2) <> "" And Text2(3) <> "" And Text2(3) <> "ERROR" Then
                Text2(4) = Round(Replace(Text2(3), "kg", "") * Text2(2), 2)
            End If
            
        Case 3
            If Text2(1) <> "" And Text2(2) <> "" And Text2(3) <> "" And Text2(3) <> "ERROR" Then
                Text2(4) = Round(Replace(Text2(3), "kg", "") * Text2(2), 2)
                TPeso(2) = Replace(Text2(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(2) = 0
            End If
            
        Case 4
            TSubtotal(2) = Text2(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text3_Change(Index As Integer)

    On Error Resume Next

    Select Case Index
    
        Case 1
            If Text3(1).Text <> "" Then
                Text3(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text3(i).Text = ""
                    Text3(i).Enabled = False
                Next
                
                Command3(2).Enabled = False
                
                TSubtotal(3) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(3) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If

        Case 2
            If Text3(1) <> "" And Text3(2) <> "" And Text3(3) <> "" And Text3(3) <> "ERROR" Then
                Text3(4) = Round(Replace(Text3(3), "kg", "") * Text3(2), 2)
            End If
            
        Case 3
            If Text3(1) <> "" And Text3(2) <> "" And Text3(3) <> "" And Text3(3) <> "ERROR" Then
                Text3(4) = Round(Replace(Text3(3), "kg", "") * Text3(2), 2)
                TPeso(3) = Replace(Text3(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(3) = 0
            End If
            
        Case 4
            TSubtotal(3) = Text3(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text4_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text4(1).Text <> "" Then
                Text4(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text4(i).Text = ""
                    Text4(i).Enabled = False
                Next
                
                Command4(2).Enabled = False
                
                TSubtotal(4) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(4) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text4(1) <> "" And Text4(2) <> "" And Text4(3) <> "" And Text4(3) <> "ERROR" Then
                Text4(4) = Round(Replace(Text4(3), "kg", "") * Text4(2), 2)
            End If
            
        Case 3
            If Text4(1) <> "" And Text4(2) <> "" And Text4(3) <> "" And Text4(3) <> "ERROR" Then
                Text4(4) = Round(Replace(Text4(3), "kg", "") * Text4(2), 2)
                TPeso(4) = Replace(Text4(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(4) = 0
            End If
            
        Case 4
            TSubtotal(4) = Text4(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text5_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text5(1).Text <> "" Then
                Text5(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text5(i).Text = ""
                    Text5(i).Enabled = False
                Next
                
                Command5(2).Enabled = False
                
                TSubtotal(5) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(5) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text5(1) <> "" And Text5(2) <> "" And Text5(3) <> "" And Text5(3) <> "ERROR" Then
                Text5(4) = Round(Replace(Text5(3), "kg", "") * Text5(2), 2)
            End If
            
        Case 3
            If Text5(1) <> "" And Text5(2) <> "" And Text5(3) <> "" And Text5(3) <> "ERROR" Then
                Text5(4) = Round(Replace(Text5(3), "kg", "") * Text5(2), 2)
                TPeso(5) = Replace(Text5(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(5) = 0
            End If
            
        Case 4
            TSubtotal(5) = Text5(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text6_Change(Index As Integer)

    On Error Resume Next

    Select Case Index
    
        Case 1
            If Text6(1).Text <> "" Then
                Text6(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text6(i).Text = ""
                    Text6(i).Enabled = False
                Next
                
                Command6(2).Enabled = False
                
                TSubtotal(6) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(6) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If

        Case 2
            If Text6(1) <> "" And Text6(2) <> "" And Text6(3) <> "" And Text6(3) <> "ERROR" Then
                Text6(4) = Round(Replace(Text6(3), "kg", "") * Text6(2), 2)
            End If
            
        Case 3
            If Text6(1) <> "" And Text6(2) <> "" And Text6(3) <> "" And Text6(3) <> "ERROR" Then
                Text6(4) = Round(Replace(Text6(3), "kg", "") * Text6(2), 2)
                TPeso(6) = Replace(Text6(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(6) = 0
            End If
            
        Case 4
            TSubtotal(6) = Text6(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text7_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text7(1).Text <> "" Then
                Text7(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text7(i).Text = ""
                    Text7(i).Enabled = False
                Next
                
                Command7(2).Enabled = False
                
                TSubtotal(7) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(7) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text7(1) <> "" And Text7(2) <> "" And Text7(3) <> "" And Text7(3) <> "ERROR" Then
                Text7(4) = Round(Replace(Text7(3), "kg", "") * Text7(2), 2)
            End If
            
        Case 3
            If Text7(1) <> "" And Text7(2) <> "" And Text7(3) <> "" And Text7(3) <> "ERROR" Then
                Text7(4) = Round(Replace(Text7(3), "kg", "") * Text7(2), 2)
                TPeso(7) = Replace(Text7(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(7) = 0
            End If
            
        Case 4
            TSubtotal(7) = Text7(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text8_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text8(1).Text <> "" Then
                Text8(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text8(i).Text = ""
                    Text8(i).Enabled = False
                Next
                
                Command8(2).Enabled = False
                
                TSubtotal(8) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(8) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text8(1) <> "" And Text8(2) <> "" And Text8(3) <> "" And Text8(3) <> "ERROR" Then
                Text8(4) = Round(Replace(Text8(3), "kg", "") * Text8(2), 2)
            End If
            
        Case 3
            If Text8(1) <> "" And Text8(2) <> "" And Text8(3) <> "" And Text8(3) <> "ERROR" Then
                Text8(4) = Round(Replace(Text8(3), "kg", "") * Text8(2), 2)
                TPeso(8) = Replace(Text8(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(8) = 0
            End If
            
        Case 4
            TSubtotal(8) = Text8(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text9_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text9(1).Text <> "" Then
                Text9(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text9(i).Text = ""
                    Text9(i).Enabled = False
                Next
                
                Command9(2).Enabled = False
                
                TSubtotal(9) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(9) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text9(1) <> "" And Text9(2) <> "" And Text9(3) <> "" And Text9(3) <> "ERROR" Then
                Text9(4) = Round(Replace(Text9(3), "kg", "") * Text9(2), 2)
            End If
            
        Case 3
            If Text9(1) <> "" And Text9(2) <> "" And Text9(3) <> "" And Text9(3) <> "ERROR" Then
                Text9(4) = Round(Replace(Text9(3), "kg", "") * Text9(2), 2)
                TPeso(9) = Replace(Text9(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(9) = 0
            End If
            
        Case 4
            TSubtotal(9) = Text9(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text10_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text10(1).Text <> "" Then
                Text10(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text10(i).Text = ""
                    Text10(i).Enabled = False
                Next
                
                Command10(2).Enabled = False
                
                TSubtotal(10) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(10) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text10(1) <> "" And Text10(2) <> "" And Text10(3) <> "" And Text10(3) <> "ERROR" Then
                Text10(4) = Round(Replace(Text10(3), "kg", "") * Text10(2), 2)
            End If
            
        Case 3
            If Text10(1) <> "" And Text10(2) <> "" And Text10(3) <> "" And Text10(3) <> "ERROR" Then
                Text10(4) = Round(Replace(Text10(3), "kg", "") * Text10(2), 2)
                TPeso(10) = Replace(Text10(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(10) = 0
            End If
            
        Case 4
            TSubtotal(10) = Text10(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text11_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text11(1).Text <> "" Then
                Text11(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text11(i).Text = ""
                    Text11(i).Enabled = False
                Next
                
                Command11(2).Enabled = False
                
                TSubtotal(11) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(11) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text11(1) <> "" And Text11(2) <> "" And Text11(3) <> "" And Text11(3) <> "ERROR" Then
                Text11(4) = Round(Replace(Text11(3), "kg", "") * Text11(2), 2)
            End If
            
        Case 3
            If Text11(1) <> "" And Text11(2) <> "" And Text11(3) <> "" And Text11(3) <> "ERROR" Then
                Text11(4) = Round(Replace(Text11(3), "kg", "") * Text11(2), 2)
                TPeso(11) = Replace(Text11(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(11) = 0
            End If
            
        Case 4
            TSubtotal(11) = Text11(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text12_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text12(1).Text <> "" Then
                Text12(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text12(i).Text = ""
                    Text12(i).Enabled = False
                Next
                
                Command12(2).Enabled = False
                
                TSubtotal(12) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(12) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text12(1) <> "" And Text12(2) <> "" And Text12(3) <> "" And Text12(3) <> "ERROR" Then
                Text12(4) = Round(Replace(Text12(3), "kg", "") * Text12(2), 2)
            End If
            
        Case 3
            If Text12(1) <> "" And Text12(2) <> "" And Text12(3) <> "" And Text12(3) <> "ERROR" Then
                Text12(4) = Round(Replace(Text12(3), "kg", "") * Text12(2), 2)
                TPeso(12) = Replace(Text12(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(12) = 0
            End If
            
        Case 4
            TSubtotal(12) = Text12(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text13_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text13(1).Text <> "" Then
                Text13(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text13(i).Text = ""
                    Text13(i).Enabled = False
                Next
                
                Command13(2).Enabled = False
                
                TSubtotal(13) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(13) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text13(1) <> "" And Text13(2) <> "" And Text13(3) <> "" And Text13(3) <> "ERROR" Then
                Text13(4) = Round(Replace(Text13(3), "kg", "") * Text13(2), 2)
            End If
            
        Case 3
            If Text13(1) <> "" And Text13(2) <> "" And Text13(3) <> "" And Text13(3) <> "ERROR" Then
                Text13(4) = Round(Replace(Text13(3), "kg", "") * Text13(2), 2)
                TPeso(13) = Replace(Text13(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(13) = 0
            End If
            
        Case 4
            TSubtotal(13) = Text13(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text14_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text14(1).Text <> "" Then
                Text14(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text14(i).Text = ""
                    Text14(i).Enabled = False
                Next
                
                Command14(2).Enabled = False
                
                TSubtotal(14) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(14) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text14(1) <> "" And Text14(2) <> "" And Text14(3) <> "" And Text14(3) <> "ERROR" Then
                Text14(4) = Round(Replace(Text14(3), "kg", "") * Text14(2), 2)
            End If
            
        Case 3
            If Text14(1) <> "" And Text14(2) <> "" And Text14(3) <> "" And Text14(3) <> "ERROR" Then
                Text14(4) = Round(Replace(Text14(3), "kg", "") * Text14(2), 2)
                TPeso(14) = Replace(Text14(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(14) = 0
            End If
            
        Case 4
            TSubtotal(14) = Text14(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub

Private Sub Text15_Change(Index As Integer)

    On Error Resume Next

    Select Case Index

        Case 1
            If Text15(1).Text <> "" Then
                Text15(2).Enabled = True
            Else
                
                For i = 2 To 4
                    Text15(i).Text = ""
                    Text15(i).Enabled = False
                Next
                
                Command15(2).Enabled = False
                
                TSubtotal(15) = 0
                Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
                TPeso(15) = 0
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            End If
        
        Case 2
            If Text15(1) <> "" And Text15(2) <> "" And Text15(3) <> "" And Text15(3) <> "ERROR" Then
                Text15(4) = Round(Replace(Text15(3), "kg", "") * Text15(2), 2)
            End If
            
        Case 3
            If Text15(1) <> "" And Text15(2) <> "" And Text15(3) <> "" And Text15(3) <> "ERROR" Then
                Text15(4) = Round(Replace(Text15(3), "kg", "") * Text15(2), 2)
                TPeso(15) = Replace(Text15(3), "kg", "")
                Text16(1) = TPeso(1) + TPeso(2) + TPeso(3) + TPeso(4) + TPeso(5) + TPeso(6) + TPeso(7) + TPeso(8) + TPeso(9) + TPeso(10) + TPeso(11) + TPeso(12) + TPeso(13) + TPeso(14) + TPeso(15)
            Else
                TPeso(15) = 0
            End If
            
        Case 4
            TSubtotal(15) = Text15(4)
            Text16(2) = TSubtotal(1) + TSubtotal(2) + TSubtotal(3) + TSubtotal(4) + TSubtotal(5) + TSubtotal(6) + TSubtotal(7) + TSubtotal(8) + TSubtotal(9) + TSubtotal(10) + TSubtotal(11) + TSubtotal(12) + TSubtotal(13) + TSubtotal(14) + TSubtotal(15)
            
    End Select

End Sub
