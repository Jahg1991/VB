VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresiòn"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpresion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1920
      TabIndex        =   70
      Top             =   360
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   11
      TabsPerRow      =   6
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Somatometrìa"
      TabPicture(0)   =   "frmImpresion.frx":324A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text2(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text2(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Audiometrìa"
      TabPicture(1)   =   "frmImpresion.frx":3266
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text4"
      Tab(1).Control(1)=   "Check2(1)"
      Tab(1).Control(2)=   "Check2(0)"
      Tab(1).Control(3)=   "Label3(14)"
      Tab(1).Control(4)=   "Label3(13)"
      Tab(1).Control(5)=   "Label3(12)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Cardiologìa"
      TabPicture(2)   =   "frmImpresion.frx":3282
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text5"
      Tab(2).Control(1)=   "Check3"
      Tab(2).Control(2)=   "Label3(16)"
      Tab(2).Control(3)=   "Label3(15)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Dental"
      TabPicture(3)   =   "frmImpresion.frx":329E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text6"
      Tab(3).Control(1)=   "Check4(1)"
      Tab(3).Control(2)=   "Check4(0)"
      Tab(3).Control(3)=   "Label3(19)"
      Tab(3).Control(4)=   "Label3(18)"
      Tab(3).Control(5)=   "Label3(17)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "D. O. C. M."
      TabPicture(4)   =   "frmImpresion.frx":32BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text7"
      Tab(4).Control(1)=   "Check5"
      Tab(4).Control(2)=   "Label3(23)"
      Tab(4).Control(3)=   "Label3(22)"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "D. O. C. C. U."
      TabPicture(5)   =   "frmImpresion.frx":32D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text8"
      Tab(5).Control(1)=   "Check6"
      Tab(5).Control(2)=   "Label3(21)"
      Tab(5).Control(3)=   "Label3(20)"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Mastografìa"
      TabPicture(6)   =   "frmImpresion.frx":32F2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text9"
      Tab(6).Control(1)=   "Check7"
      Tab(6).Control(2)=   "Label3(25)"
      Tab(6).Control(3)=   "Label3(24)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Laboratorio"
      TabPicture(7)   =   "frmImpresion.frx":330E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Text3(3)"
      Tab(7).Control(1)=   "Text3(2)"
      Tab(7).Control(2)=   "Text3(1)"
      Tab(7).Control(3)=   "Text3(0)"
      Tab(7).Control(4)=   "Label3(11)"
      Tab(7).Control(5)=   "Label3(10)"
      Tab(7).Control(6)=   "Label3(9)"
      Tab(7).Control(7)=   "Label3(8)"
      Tab(7).ControlCount=   8
      TabCaption(8)   =   "Nutriciòn"
      TabPicture(8)   =   "frmImpresion.frx":332A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Check8(1)"
      Tab(8).Control(1)=   "Text10"
      Tab(8).Control(2)=   "Check8(0)"
      Tab(8).Control(3)=   "Label3(28)"
      Tab(8).Control(4)=   "Label3(27)"
      Tab(8).Control(5)=   "Label3(26)"
      Tab(8).ControlCount=   6
      TabCaption(9)   =   "Optometrìa"
      TabPicture(9)   =   "frmImpresion.frx":3346
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Text11"
      Tab(9).Control(1)=   "Check9"
      Tab(9).Control(2)=   "Label3(30)"
      Tab(9).Control(3)=   "Label3(29)"
      Tab(9).ControlCount=   4
      TabCaption(10)  =   "Tuberculosis"
      TabPicture(10)  =   "frmImpresion.frx":3362
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Text12"
      Tab(10).Control(1)=   "Check10"
      Tab(10).Control(2)=   "Label3(32)"
      Tab(10).Control(3)=   "Label3(31)"
      Tab(10).ControlCount=   4
      Begin VB.CheckBox Check8 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -72840
         TabIndex        =   69
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   68
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   67
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   66
         Top             =   1560
         Width           =   6735
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   65
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   64
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   63
         Top             =   1200
         Width           =   6735
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check3"
         Height          =   255
         Left            =   -72840
         TabIndex        =   62
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72840
         TabIndex        =   61
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   -72840
         TabIndex        =   60
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72840
         TabIndex        =   59
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72840
         TabIndex        =   58
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72840
         TabIndex        =   57
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   56
         Top             =   1560
         Width           =   6735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -72840
         TabIndex        =   55
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   -72840
         TabIndex        =   54
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   53
         Top             =   1200
         Width           =   6735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -72840
         TabIndex        =   52
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72840
         TabIndex        =   51
         Top             =   1560
         Width           =   6735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -72840
         TabIndex        =   50
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   -72840
         TabIndex        =   49
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   -72840
         TabIndex        =   27
         Top             =   1920
         Width           =   6735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   -72840
         TabIndex        =   26
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   -72840
         TabIndex        =   25
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   -72840
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   2160
         TabIndex        =   19
         Top             =   3360
         Width           =   6735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Width           =   6735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2160
         TabIndex        =   17
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   2160
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   2160
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   2160
         TabIndex        =   14
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   -74880
         TabIndex        =   48
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   -74880
         TabIndex        =   47
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   -74880
         TabIndex        =   46
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   -74880
         TabIndex        =   45
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   -74880
         TabIndex        =   44
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Platica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   -74880
         TabIndex        =   43
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   -74880
         TabIndex        =   42
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   -74880
         TabIndex        =   41
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   -74880
         TabIndex        =   40
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   -74880
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   -74880
         TabIndex        =   38
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   -74880
         TabIndex        =   37
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   -74880
         TabIndex        =   36
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   -74880
         TabIndex        =   35
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Revisiòn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   -74880
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Limpieza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   -74880
         TabIndex        =   33
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   -74760
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asistencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -74880
         TabIndex        =   31
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -74880
         TabIndex        =   30
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Prueba de audiciòn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -74880
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lavado de oìdos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -74880
         TabIndex        =   28
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   23
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Glucosa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   22
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Trigliceridos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   21
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Colesterol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   20
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Otras Vacunas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vacuna Toxoide"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tensiòn arterial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gènero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de nacimeinto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3840
      Picture         =   "frmImpresion.frx":337E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   13
      Left            =   8880
      Picture         =   "frmImpresion.frx":B101
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   120
      Picture         =   "frmImpresion.frx":B688
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            VarNombre = Text1(0).Text
            VarFecha_nacimiento = Text2(0).Text
            VarGenero = Text2(1).Text
            VarPeso = Text2(2).Text
            VarTalla = Text2(3).Text
            VarTension_arterial = Text2(4).Text
            If Check1.Value = 1 Then
                VarVacuna_toxoide = "Si"
            Else
                VarVacuna_toxoide = "No"
            End If
            VarOtras_vacunas = Text2(5).Text
            VarObservaciones_somatometria = Text2(6).Text
            VarColesterol = Text3(0).Text
            VarTrigliceridos = Text3(1).Text
            VarGlucosa = Text3(2).Text
            VarObservaciones_laboratorio = Text3(3).Text
            If Check2(0).Value = 1 Then
                VarLavado_oidos = "Si"
            Else
                VarLavado_oidos = "No"
            End If
            If Check2(1).Value = 1 Then
                VarPrueba_audicion = "Si"
            Else
                VarPrueba_audicion = "No"
            End If
            VarObservaciones_audiometria = Text4.Text
            If Check3.Value = 1 Then
                VarCardiologia = "Asistencia " & Text5.Text
            Else
                VarCardiologia = ""
            End If
            If Check4(0).Value = 1 Then
                VarLimpieza_dental = "Si"
            Else
                VarLimpieza_dental = "No"
            End If
            If Check4(1).Value = 1 Then
                VarRevision_dental = "Si"
            Else
                VarRevision_dental = "No"
            End If
            If Check6.Value = 1 Then
                VarDoccu = "Asistencia " & Text8.Text
            Else
                VarDoccu = ""
            End If
            If Check5.Value = 1 Then
                VarDocm = "Asistencia " & Text7.Text
            Else
                VarDocm = ""
            End If
            If Check7.Value = 1 Then
                Varmastografia = "Asistencia " & Text9.Text
            Else
                Varmastografia = ""
            End If
            If Check8(0).Value = 1 Then
                VarConsulta_nutricion = "Si"
            Else
                VarConsulta_nutricion = "No"
            End If
            If Check8(1).Value = 1 Then
                VarPlatica_nutricion = "Si"
            Else
                VarLimpieza_dental = "No"
            End If
            VarObservaciones_nutricion = Text10.Text
            If Check9.Value = 1 Then
                VarObservaciones_optometria = "Asistencia " & Text11.Text
            Else
                VarObservaciones_optometria = ""
            End If
            If Check10.Value = 1 Then
                VarObservaciones_tuberculosis = "Asistencia " & Text12.Text
            Else
                VarObservaciones_tuberculosis = ""
            End If
            Set DataReport1.DataSource = Rs
            DataReport1.Sections("Sección2").Controls("Etiqueta41").Caption = VarNombre
            DataReport1.Sections("Sección2").Controls("Etiqueta42").Caption = VarFecha_nacimiento
            DataReport1.Sections("Sección2").Controls("Etiqueta43").Caption = VarGenero
            DataReport1.Sections("Sección1").Controls("Etiqueta44").Caption = VarPeso
            DataReport1.Sections("Sección1").Controls("Etiqueta45").Caption = VarTalla
            DataReport1.Sections("Sección1").Controls("Etiqueta46").Caption = VarTension_arterial
            DataReport1.Sections("Sección1").Controls("Etiqueta47").Caption = VarVacuna_toxoide
            DataReport1.Sections("Sección1").Controls("Etiqueta48").Caption = VarOtras_vacunas
            DataReport1.Sections("Sección1").Controls("Etiqueta68").Caption = VarObservaciones_somatometria
            DataReport1.Sections("Sección1").Controls("Etiqueta58").Caption = VarColesterol
            DataReport1.Sections("Sección1").Controls("Etiqueta59").Caption = VarTrigliceridos
            DataReport1.Sections("Sección1").Controls("Etiqueta60").Caption = VarGlucosa
            DataReport1.Sections("Sección1").Controls("Etiqueta61").Caption = VarObservaciones_laboratorio
            DataReport1.Sections("Sección1").Controls("Etiqueta49").Caption = VarLavado_oidos
            DataReport1.Sections("Sección1").Controls("Etiqueta50").Caption = VarPrueba_audicion
            DataReport1.Sections("Sección1").Controls("Etiqueta51").Caption = VarObservaciones_audiometria
            DataReport1.Sections("Sección1").Controls("Etiqueta52").Caption = VarCardiologia
            DataReport1.Sections("Sección1").Controls("Etiqueta53").Caption = VarLimpieza_dental
            DataReport1.Sections("Sección1").Controls("Etiqueta54").Caption = VarRevision_dental
            DataReport1.Sections("Sección1").Controls("Etiqueta55").Caption = VarObservaciones_dental
            DataReport1.Sections("Sección1").Controls("Etiqueta57").Caption = VarDoccu
            DataReport1.Sections("Sección1").Controls("Etiqueta56").Caption = VarDocm
            DataReport1.Sections("Sección1").Controls("Etiqueta62").Caption = Varmastografia
            DataReport1.Sections("Sección1").Controls("Etiqueta63").Caption = VarConsulta_nutricion
            DataReport1.Sections("Sección1").Controls("Etiqueta64").Caption = VarPlatica_nutricion
            DataReport1.Sections("Sección1").Controls("Etiqueta65").Caption = VarObservaciones_nutricion
            DataReport1.Sections("Sección1").Controls("Etiqueta66").Caption = VarObservaciones_optometria
            DataReport1.Sections("Sección1").Controls("Etiqueta67").Caption = VarObservaciones_tuberculosis
            Form4.Show
            Form1.Enabled = False
    End Select
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 13
            Form2.Show
            Form1.Enabled = False
    End Select
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
        Case 1
            On Error Resume Next
            With Rs
                .Requery
                If option1.Value = True Then
                    .Filter = "Id like '" & Text1(1) & "'"
                End If
            End With
    End Select
End Sub
