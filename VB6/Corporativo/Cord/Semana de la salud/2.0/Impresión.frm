VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FormImpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de resultados"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "Impresión.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   970
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Somatometría"
      TabPicture(0)   =   "Impresión.frx":324A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Laboratorio"
      TabPicture(1)   =   "Impresión.frx":3266
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dental"
      TabPicture(2)   =   "Impresión.frx":3282
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Nutrición"
      TabPicture(3)   =   "Impresión.frx":329E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Salud de  la  mujer"
      TabPicture(4)   =   "Impresión.frx":32BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Optometría"
      TabPicture(5)   =   "Impresión.frx":32D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Audiometría"
      TabPicture(6)   =   "Impresión.frx":32F2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Tuberculosis"
      TabPicture(7)   =   "Impresión.frx":330E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame8"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Cardiología"
      TabPicture(8)   =   "Impresión.frx":332A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame9"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   71
         Top             =   600
         Width           =   10095
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   1
            Left            =   1920
            TabIndex        =   78
            Top             =   720
            Width           =   8055
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   77
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Asistencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   480
            TabIndex        =   73
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   68
         Top             =   600
         Width           =   10095
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   1
            Left            =   1920
            TabIndex        =   76
            Top             =   720
            Width           =   8055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   75
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Asistencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   480
            TabIndex        =   69
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   64
         Top             =   600
         Width           =   10095
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   1
            Left            =   1920
            TabIndex        =   74
            Top             =   720
            Width           =   8055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Asistencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   67
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   65
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   57
         Top             =   600
         Width           =   10095
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   2
            Left            =   1920
            TabIndex        =   61
            Top             =   720
            Width           =   8055
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   60
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Asistencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   480
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   48
         Top             =   600
         Width           =   10095
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Index           =   3
            Left            =   1920
            TabIndex        =   56
            Top             =   1440
            Width           =   8055
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   55
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   54
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   53
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "DOCMA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   720
            TabIndex        =   52
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "DOCCU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   840
            TabIndex        =   50
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Mastografía"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   480
            TabIndex        =   49
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   10095
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Index           =   2
            Left            =   1920
            TabIndex        =   47
            Top             =   960
            Width           =   8175
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   46
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   45
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Asistencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   720
            TabIndex        =   44
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   3840
            TabIndex        =   43
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   3840
            TabIndex        =   42
            Top             =   600
            Width           =   6375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   1320
            TabIndex        =   40
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   34
         Top             =   600
         Width           =   10095
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Index           =   1
            Left            =   1800
            TabIndex        =   38
            Top             =   720
            Width           =   8175
         End
         Begin VB.Label Label4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Asistencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   36
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74880
         TabIndex        =   23
         Top             =   600
         Width           =   10095
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Index           =   4
            Left            =   1920
            TabIndex        =   33
            Top             =   1800
            Width           =   8055
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   32
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   31
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   30
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   29
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Colesterol"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Trigliceridos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Glucosa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "PSA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   25
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   10095
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   9
            Left            =   2520
            TabIndex        =   20
            Top             =   2760
            Width           =   7455
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2520
            TabIndex        =   19
            Top             =   2400
            Width           =   7455
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2520
            TabIndex        =   18
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2520
            TabIndex        =   17
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2520
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   15
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2520
            TabIndex        =   14
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha de nacimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Género"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   11
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Peso en Kg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   1080
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Talla en cm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1200
            TabIndex        =   9
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tensión arterial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   8
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Vacuna toxoide"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   720
            TabIndex        =   7
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Otras vacunas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   840
            TabIndex        =   6
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   720
            TabIndex        =   5
            Top             =   2760
            Width           =   1695
         End
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
      DefColWidth     =   267
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   4200
      TabIndex        =   63
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   1560
      TabIndex        =   62
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   22
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Id Asistente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre Completo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu Cancelar 
         Caption         =   "Cancelar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FormImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    On Error Resume Next
    FormImpresion.Text1(1).Text = ""
    FormImpresion.Label2(2).Caption = ""
    FormImpresion.Label2(3).Caption = ""
    FormImpresion.Label2(4).Caption = ""
    FormImpresion.Label2(5).Caption = ""
    FormImpresion.Label2(6).Caption = ""
    FormImpresion.Label2(7).Caption = ""
    FormImpresion.Label2(8).Caption = ""
    FormImpresion.Label2(9).Caption = ""
    FormImpresion.Label3(0).Caption = ""
    FormImpresion.Label3(1).Caption = ""
    FormImpresion.Label3(2).Caption = ""
    FormImpresion.Label3(3).Caption = ""
    FormImpresion.Label3(4).Caption = ""
    FormImpresion.Label4(0).Caption = ""
    FormImpresion.Label4(1).Caption = ""
    FormImpresion.Label5(0).Caption = ""
    FormImpresion.Label5(1).Caption = ""
    FormImpresion.Label5(2).Caption = ""
    FormImpresion.Label6(0).Caption = ""
    FormImpresion.Label6(1).Caption = ""
    FormImpresion.Label6(2).Caption = ""
    FormImpresion.Label6(3).Caption = ""
    FormImpresion.Label7(1).Caption = ""
    FormImpresion.Label7(2).Caption = ""
    FormImpresion.Label8(0).Caption = ""
    FormImpresion.Label8(1).Caption = ""
    FormImpresion.Label9(0).Caption = ""
    FormImpresion.Label9(1).Caption = ""
    FormImpresion.Label10(0).Caption = ""
    FormImpresion.Label10(1).Caption = ""
    With RsNombre
        .Filter = ""
        .MoveFirst
    End With
    Set DataGrid1.DataSource = RsNombre
    Text1(1).SetFocus
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RsAllDate
        If .State = 1 Then .Close
            .Open "Select * from All_date", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsNombre
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set DataGrid1.DataSource = RsNombre
    Set FormImpresion.Label1(15).DataSource = RsNombre
    Set FormImpresion.Label1(16).DataSource = RsNombre
    Set FormImpresion.Label2(2).DataSource = RsAllDate
    Set FormImpresion.Label2(3).DataSource = RsAllDate
    Set FormImpresion.Label2(4).DataSource = RsAllDate
    Set FormImpresion.Label2(5).DataSource = RsAllDate
    Set FormImpresion.Label2(6).DataSource = RsAllDate
    Set FormImpresion.Label2(7).DataSource = RsAllDate
    Set FormImpresion.Label2(8).DataSource = RsAllDate
    Set FormImpresion.Label2(9).DataSource = RsAllDate
    Set FormImpresion.Label3(0).DataSource = RsAllDate
    Set FormImpresion.Label3(1).DataSource = RsAllDate
    Set FormImpresion.Label3(2).DataSource = RsAllDate
    Set FormImpresion.Label3(3).DataSource = RsAllDate
    Set FormImpresion.Label3(4).DataSource = RsAllDate
    Set FormImpresion.Label4(0).DataSource = RsAllDate
    Set FormImpresion.Label4(1).DataSource = RsAllDate
    Set FormImpresion.Label5(0).DataSource = RsAllDate
    Set FormImpresion.Label5(1).DataSource = RsAllDate
    Set FormImpresion.Label5(2).DataSource = RsAllDate
    Set FormImpresion.Label6(0).DataSource = RsAllDate
    Set FormImpresion.Label6(1).DataSource = RsAllDate
    Set FormImpresion.Label6(2).DataSource = RsAllDate
    Set FormImpresion.Label6(3).DataSource = RsAllDate
    Set FormImpresion.Label7(1).DataSource = RsAllDate
    Set FormImpresion.Label7(2).DataSource = RsAllDate
    Set FormImpresion.Label8(0).DataSource = RsAllDate
    Set FormImpresion.Label8(1).DataSource = RsAllDate
    Set FormImpresion.Label9(0).DataSource = RsAllDate
    Set FormImpresion.Label9(1).DataSource = RsAllDate
    Set FormImpresion.Label10(0).DataSource = RsAllDate
    Set FormImpresion.Label10(1).DataSource = RsAllDate
    FormImpresion.Label1(15).DataField = ("ID_AST")
    FormImpresion.Label1(16).DataField = ("NOMBRE")
    FormImpresion.Label2(2).DataField = ("FECHA_NACIMIENTO")
    FormImpresion.Label2(3).DataField = ("GENERO")
    FormImpresion.Label2(4).DataField = ("PESO")
    FormImpresion.Label2(5).DataField = ("TALLA")
    FormImpresion.Label2(6).DataField = ("TA")
    FormImpresion.Label2(7).DataField = ("VACUNA_TOXOIDE")
    FormImpresion.Label2(8).DataField = ("OTRAS_VACUNAS")
    FormImpresion.Label2(9).DataField = ("OBSERVACIONES_SOMATOMETRÍA")
    FormImpresion.Label3(0).DataField = ("COLESTEROL")
    FormImpresion.Label3(1).DataField = ("TRIGLICERIDOS")
    FormImpresion.Label3(2).DataField = ("GLUCOSA")
    FormImpresion.Label3(3).DataField = ("PSA")
    FormImpresion.Label3(4).DataField = ("OBSERVACIONES_LABORATORIO")
    FormImpresion.Label4(0).DataField = ("ASISTENCIA_DENTAL")
    FormImpresion.Label4(1).DataField = ("OBSERVACIONES_DENTAL")
    FormImpresion.Label5(0).DataField = ("ASISTENCIA_NUTRICION")
    FormImpresion.Label5(1).DataField = ("TIPO")
    FormImpresion.Label5(2).DataField = ("OBSERVACIONES_NUTRICION")
    FormImpresion.Label6(0).DataField = ("DOCMA")
    FormImpresion.Label6(1).DataField = ("DOCCU")
    FormImpresion.Label6(2).DataField = ("MASTOGRAFIA")
    FormImpresion.Label6(3).DataField = ("OBSERVACIONES_SALUD_DE_LA_MUJER")
    FormImpresion.Label7(1).DataField = ("OPTOMETRIA")
    FormImpresion.Label7(2).DataField = ("OBSERVACIONES_OPTOMETRIA")
    FormImpresion.Label8(0).DataField = ("ASISTENCIA_AUDIOMETRIA")
    FormImpresion.Label8(1).DataField = ("OBSERVACIONES_AUDIOMETRIA")
    FormImpresion.Label9(0).DataField = ("ASISTENCIA_TUBERCULOSIS")
    FormImpresion.Label9(1).DataField = ("OBSERVACIONES_TUBERCULOSIS")
    FormImpresion.Label10(0).DataField = ("ASISTENCIA_CARDIOLOGIA")
    FormImpresion.Label10(1).DataField = ("OBSERVACIONES_CARDIOLOGIA")
    Text1(1).SetFocus
End Sub
Private Sub Imprimir_Click()
    On Error Resume Next
    RsNombre.Filter = "NOMBRE LIKE '*" & Label1(16) & "*'"
    Set DataReport1.DataSource = RsNombre
    DRID_AST = FormImpresion.Label1(15).Caption
    DRNOMBRE = FormImpresion.Label1(16).Caption
    DRFECHA_NACIMIENTO = FormImpresion.Label2(2).Caption
    DRGENERO = FormImpresion.Label2(3).Caption
    DRPESO = FormImpresion.Label2(4).Caption
    DRTALLA = FormImpresion.Label2(5).Caption
    DRTA = FormImpresion.Label2(6).Caption
    DRVACUNA_TOXOIDE = FormImpresion.Label2(7).Caption
    DROTRAS_VACUNAS = FormImpresion.Label2(8).Caption
    DROBSERVACIONES_SOMATOMETRÍA = FormImpresion.Label2(9).Caption
    DRCOLESTEROL = FormImpresion.Label3(0).Caption
    DRTRIGLICERIDOS = FormImpresion.Label3(1).Caption
    DRGLUCOSA = FormImpresion.Label3(2).Caption
    DRPSA = FormImpresion.Label3(3).Caption
    DROBSERVACIONES_LABORATORIO = FormImpresion.Label3(4).Caption
    DRASISTENCIA_DENTAL = FormImpresion.Label4(0).Caption
    DROBSERVACIONES_DENTAL = FormImpresion.Label4(1).Caption
    DRASISTENCIA_NUTRICION = FormImpresion.Label5(0).Caption
    DRTIPO = FormImpresion.Label5(1).Caption
    DROBSERVACIONES_NUTRICION = FormImpresion.Label5(2).Caption
    DRDOCMA = FormImpresion.Label6(0).Caption
    DRDOCCU = FormImpresion.Label6(1).Caption
    DRMASTOGRAFIA = FormImpresion.Label6(2).Caption
    DROBSERVACIONES_SALUD_DE_LA_MUJER = FormImpresion.Label6(3).Caption
    DROPTOMETRIA = FormImpresion.Label7(1).Caption
    DROBSERVACIONES_OPTOMETRIA = FormImpresion.Label7(2).Caption
    DRASISTENCIA_AUDIOMETRIA = Label8(0).Caption
    DROBSERVACIONES_AUDIOMETRIA = Label8(1).Caption
    DRASISTENCIA_TUBERCULOSIS = Label9(0).Caption
    DROBSERVACIONES_TUBERCULOSIS = Label9(1).Caption
    DRASISTENCIA_CARDIOLOGIA = Label10(0).Caption
    DROBSERVACIONES_CARDIOLOGIA = Label10(1).Caption
    DataReport1.Sections("Sección2").Controls("Etiqueta35").Caption = DRNOMBRE
    DataReport1.Sections("Sección1").Controls("Etiqueta36").Caption = DRFECHA_NACIMIENTO
    DataReport1.Sections("Sección1").Controls("Etiqueta37").Caption = DRGENERO
    DataReport1.Sections("Sección1").Controls("Etiqueta38").Caption = DRPESO
    DataReport1.Sections("Sección1").Controls("Etiqueta39").Caption = DRTALLA
    DataReport1.Sections("Sección1").Controls("Etiqueta40").Caption = DRTA
    DataReport1.Sections("Sección1").Controls("Etiqueta41").Caption = DRVACUNA_TOXOIDE
    DataReport1.Sections("Sección1").Controls("Etiqueta42").Caption = DROTRAS_VACUNAS
    DataReport1.Sections("Sección1").Controls("Etiqueta43").Caption = DROBSERVACIONES_SOMATOMETRÍA
    DataReport1.Sections("Sección1").Controls("Etiqueta44").Caption = DRCOLESTEROL
    DataReport1.Sections("Sección1").Controls("Etiqueta45").Caption = DRTRIGLICERIDOS
    DataReport1.Sections("Sección1").Controls("Etiqueta46").Caption = DRGLUCOSA
    DataReport1.Sections("Sección1").Controls("Etiqueta47").Caption = DRPSA
    DataReport1.Sections("Sección1").Controls("Etiqueta48").Caption = DROBSERVACIONES_LABORATORIO
    DataReport1.Sections("Sección1").Controls("Etiqueta49").Caption = DRASISTENCIA_DENTAL
    DataReport1.Sections("Sección1").Controls("Etiqueta50").Caption = DROBSERVACIONES_DENTAL
    DataReport1.Sections("Sección1").Controls("Etiqueta51").Caption = DRTIPO
    DataReport1.Sections("Sección1").Controls("Etiqueta52").Caption = DROBSERVACIONES_NUTRICION
    DataReport1.Sections("Sección1").Controls("Etiqueta53").Caption = DRDOCMA
    DataReport1.Sections("Sección1").Controls("Etiqueta54").Caption = DRDOCCU
    DataReport1.Sections("Sección1").Controls("Etiqueta55").Caption = DRMASTOGRAFIA
    DataReport1.Sections("Sección1").Controls("Etiqueta56").Caption = DROBSERVACIONES_SALUD_DE_LA_MUJER
    DataReport1.Sections("Sección1").Controls("Etiqueta58").Caption = DROPTOMETRIA
    DataReport1.Sections("Sección1").Controls("Etiqueta59").Caption = DROBSERVACIONES_OPTOMETRIA
    DataReport1.Sections("Sección1").Controls("Etiqueta67").Caption = DRASISTENCIA_AUDIOMETRIA
    DataReport1.Sections("Sección1").Controls("Etiqueta68").Caption = DROBSERVACIONES_AUDIOMETRIA
    DataReport1.Sections("Sección1").Controls("Etiqueta69").Caption = DRASISTENCIA_TUBERCULOSIS
    DataReport1.Sections("Sección1").Controls("Etiqueta70").Caption = DROBSERVACIONES_TUBERCULOSIS
    DataReport1.Sections("Sección1").Controls("Etiqueta71").Caption = DRASISTENCIA_CARDIOLOGIA
    DataReport1.Sections("Sección1").Controls("Etiqueta72").Caption = DROBSERVACIONES_CARDIOLOGIA
    DataReport1.Show
End Sub
Private Sub Label1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 15
            Set FormImpresion.Label2(2).DataSource = RsAllDate
            Set FormImpresion.Label2(3).DataSource = RsAllDate
            Set FormImpresion.Label2(4).DataSource = RsAllDate
            Set FormImpresion.Label2(5).DataSource = RsAllDate
            Set FormImpresion.Label2(6).DataSource = RsAllDate
            Set FormImpresion.Label2(7).DataSource = RsAllDate
            Set FormImpresion.Label2(8).DataSource = RsAllDate
            Set FormImpresion.Label2(9).DataSource = RsAllDate
            Set FormImpresion.Label3(0).DataSource = RsAllDate
            Set FormImpresion.Label3(1).DataSource = RsAllDate
            Set FormImpresion.Label3(2).DataSource = RsAllDate
            Set FormImpresion.Label3(3).DataSource = RsAllDate
            Set FormImpresion.Label3(4).DataSource = RsAllDate
            Set FormImpresion.Label4(0).DataSource = RsAllDate
            Set FormImpresion.Label4(1).DataSource = RsAllDate
            Set FormImpresion.Label5(0).DataSource = RsAllDate
            Set FormImpresion.Label5(1).DataSource = RsAllDate
            Set FormImpresion.Label5(2).DataSource = RsAllDate
            Set FormImpresion.Label6(0).DataSource = RsAllDate
            Set FormImpresion.Label6(1).DataSource = RsAllDate
            Set FormImpresion.Label6(2).DataSource = RsAllDate
            Set FormImpresion.Label6(3).DataSource = RsAllDate
            Set FormImpresion.Label7(1).DataSource = RsAllDate
            Set FormImpresion.Label7(2).DataSource = RsAllDate
            Set FormImpresion.Label8(0).DataSource = RsAllDate
            Set FormImpresion.Label8(1).DataSource = RsAllDate
            Set FormImpresion.Label9(0).DataSource = RsAllDate
            Set FormImpresion.Label9(1).DataSource = RsAllDate
            Set FormImpresion.Label10(0).DataSource = RsAllDate
            Set FormImpresion.Label10(1).DataSource = RsAllDate
            FormImpresion.Label2(2).DataField = ("FECHA_NACIMIENTO")
            FormImpresion.Label2(3).DataField = ("GENERO")
            FormImpresion.Label2(4).DataField = ("PESO")
            FormImpresion.Label2(5).DataField = ("TALLA")
            FormImpresion.Label2(6).DataField = ("TA")
            FormImpresion.Label2(7).DataField = ("VACUNA_TOXOIDE")
            FormImpresion.Label2(8).DataField = ("OTRAS_VACUNAS")
            FormImpresion.Label2(9).DataField = ("OBSERVACIONES_SOMATOMETRÍA")
            FormImpresion.Label3(0).DataField = ("COLESTEROL")
            FormImpresion.Label3(1).DataField = ("TRIGLICERIDOS")
            FormImpresion.Label3(2).DataField = ("GLUCOSA")
            FormImpresion.Label3(3).DataField = ("PSA")
            FormImpresion.Label3(4).DataField = ("OBSERVACIONES_LABORATORIO")
            FormImpresion.Label4(0).DataField = ("ASISTENCIA_DENTAL")
            FormImpresion.Label4(1).DataField = ("OBSERVACIONES_DENTAL")
            FormImpresion.Label5(0).DataField = ("ASISTENCIA_NUTRICION")
            FormImpresion.Label5(1).DataField = ("TIPO")
            FormImpresion.Label5(2).DataField = ("OBSERVACIONES_NUTRICION")
            FormImpresion.Label6(0).DataField = ("DOCMA")
            FormImpresion.Label6(1).DataField = ("DOCCU")
            FormImpresion.Label6(2).DataField = ("MASTOGRAFIA")
            FormImpresion.Label6(3).DataField = ("OBSERVACIONES_SALUD_DE_LA_MUJER")
            FormImpresion.Label7(1).DataField = ("OPTOMETRIA")
            FormImpresion.Label7(2).DataField = ("OBSERVACIONES_OPTOMETRIA")
            FormImpresion.Label8(0).DataField = ("ASISTENCIA_AUDIOMETRIA")
            FormImpresion.Label8(1).DataField = ("OBSERVACIONES_AUDIOMETRIA")
            FormImpresion.Label9(0).DataField = ("ASISTENCIA_TUBERCULOSIS")
            FormImpresion.Label9(1).DataField = ("OBSERVACIONES_TUBERCULOSIS")
            FormImpresion.Label10(0).DataField = ("ASISTENCIA_CARDIOLOGIA")
            FormImpresion.Label10(1).DataField = ("OBSERVACIONES_CARDIOLOGIA")
            With RsAllDate
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "ID_AST LIKE '*" & Label1(15) & "*'"
                Else
                    .Filter = ""
                    .MoveFirst
                End If
            End With
    End Select
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    FormImpresion.Text1(1).Text = ""
    FormImpresion.Label2(2).Caption = ""
    FormImpresion.Label2(3).Caption = ""
    FormImpresion.Label2(4).Caption = ""
    FormImpresion.Label2(5).Caption = ""
    FormImpresion.Label2(6).Caption = ""
    FormImpresion.Label2(7).Caption = ""
    FormImpresion.Label2(8).Caption = ""
    FormImpresion.Label2(9).Caption = ""
    FormImpresion.Label3(0).Caption = ""
    FormImpresion.Label3(1).Caption = ""
    FormImpresion.Label3(2).Caption = ""
    FormImpresion.Label3(3).Caption = ""
    FormImpresion.Label3(4).Caption = ""
    FormImpresion.Label4(0).Caption = ""
    FormImpresion.Label4(1).Caption = ""
    FormImpresion.Label5(0).Caption = ""
    FormImpresion.Label5(1).Caption = ""
    FormImpresion.Label5(2).Caption = ""
    FormImpresion.Label6(0).Caption = ""
    FormImpresion.Label6(1).Caption = ""
    FormImpresion.Label6(2).Caption = ""
    FormImpresion.Label6(3).Caption = ""
    FormImpresion.Label7(1).Caption = ""
    FormImpresion.Label7(2).Caption = ""
    FormImpresion.Label8(0).Caption = ""
    FormImpresion.Label8(1).Caption = ""
    FormImpresion.Label9(0).Caption = ""
    FormImpresion.Label9(1).Caption = ""
    FormImpresion.Label10(0).Caption = ""
    FormImpresion.Label10(1).Caption = ""
    Unload Me
    Form1.Enabled = True
End Sub
Private Sub Text1_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 1
        Set FormImpresion.Label1(15).DataSource = RsNombre
        Set FormImpresion.Label1(16).DataSource = RsNombre
        FormImpresion.Label1(15).DataField = ("ID_AST")
        FormImpresion.Label1(16).DataField = ("NOMBRE")
        With RsNombre
            .Requery
            If OPTION1.Value = True Then
                .Filter = "NOMBRE LIKE '*" & Text1(1) & "*'"
            Else
                .Filter = ""
                Set DataGrid1.DataSource = RsNombre
                .MoveFirst
            End If
        End With
    End Select
End Sub
