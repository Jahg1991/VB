VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Usuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   9180
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Usuarios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nuevo"
      TabPicture(0)   =   "Usuarios.frx":324A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registrados"
      TabPicture(1)   =   "Usuarios.frx":3266
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6495
         Index           =   1
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   8655
         Begin VB.CommandButton Command3 
            Caption         =   "Eliminar"
            Height          =   495
            Index           =   4
            Left            =   7320
            TabIndex        =   76
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Último"
            Height          =   495
            Index           =   3
            Left            =   4080
            TabIndex        =   75
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Siguiente"
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   74
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Anterior"
            Height          =   495
            Index           =   1
            Left            =   1440
            TabIndex        =   73
            Top             =   5880
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Primero"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   390
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   34
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Height          =   390
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   1560
            TabIndex        =   33
            Top             =   240
            Width           =   2775
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Actualizar"
            Height          =   495
            Left            =   6000
            TabIndex        =   32
            Top             =   5880
            Width           =   1215
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   4455
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7858
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Archivo"
            TabPicture(0)   =   "Usuarios.frx":3282
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label1(28)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Check7(0)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Módulos"
            TabPicture(1)   =   "Usuarios.frx":329E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Check8(9)"
            Tab(1).Control(1)=   "Check8(8)"
            Tab(1).Control(2)=   "Check8(7)"
            Tab(1).Control(3)=   "Check8(6)"
            Tab(1).Control(4)=   "Check8(5)"
            Tab(1).Control(5)=   "Check8(4)"
            Tab(1).Control(6)=   "Check8(3)"
            Tab(1).Control(7)=   "Check8(2)"
            Tab(1).Control(8)=   "Check8(1)"
            Tab(1).Control(9)=   "Check8(0)"
            Tab(1).Control(10)=   "Label1(43)"
            Tab(1).Control(11)=   "Label1(42)"
            Tab(1).Control(12)=   "Label1(41)"
            Tab(1).Control(13)=   "Label1(35)"
            Tab(1).Control(14)=   "Label1(24)"
            Tab(1).Control(15)=   "Label1(23)"
            Tab(1).Control(16)=   "Label1(22)"
            Tab(1).Control(17)=   "Label1(21)"
            Tab(1).Control(18)=   "Label1(20)"
            Tab(1).Control(19)=   "Label1(19)"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "Edición"
            TabPicture(2)   =   "Usuarios.frx":32BA
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label1(29)"
            Tab(2).Control(1)=   "Label1(30)"
            Tab(2).Control(2)=   "Label1(31)"
            Tab(2).Control(3)=   "Label1(32)"
            Tab(2).Control(4)=   "Check9(0)"
            Tab(2).Control(5)=   "Check9(1)"
            Tab(2).Control(6)=   "Check9(2)"
            Tab(2).ControlCount=   7
            TabCaption(3)   =   "Reportes"
            TabPicture(3)   =   "Usuarios.frx":32D6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Check10(1)"
            Tab(3).Control(1)=   "Check10(0)"
            Tab(3).Control(2)=   "Label1(37)"
            Tab(3).Control(3)=   "Label1(33)"
            Tab(3).ControlCount=   4
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   9
               Left            =   -71880
               TabIndex        =   96
               Top             =   3480
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   8
               Left            =   -71880
               TabIndex        =   95
               Top             =   3120
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   7
               Left            =   -71880
               TabIndex        =   94
               Top             =   2760
               Width           =   255
            End
            Begin VB.CheckBox Check10 
               Caption         =   "Check10"
               Height          =   270
               Index           =   1
               Left            =   -73320
               TabIndex        =   84
               Top             =   960
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   6
               Left            =   -71880
               TabIndex        =   80
               Top             =   3840
               Width           =   255
            End
            Begin VB.CheckBox Check10 
               Caption         =   "Check10"
               Height          =   270
               Index           =   0
               Left            =   -73320
               TabIndex        =   71
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Check9"
               Height          =   270
               Index           =   2
               Left            =   -71640
               TabIndex        =   69
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Check9"
               Height          =   270
               Index           =   1
               Left            =   -71640
               TabIndex        =   68
               Top             =   1320
               Width           =   255
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Check9"
               Height          =   270
               Index           =   0
               Left            =   -71640
               TabIndex        =   67
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   5
               Left            =   -71880
               TabIndex        =   62
               Top             =   2400
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   4
               Left            =   -71880
               TabIndex        =   61
               Top             =   2040
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   3
               Left            =   -71880
               TabIndex        =   60
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   2
               Left            =   -71880
               TabIndex        =   59
               Top             =   1320
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   1
               Left            =   -71880
               TabIndex        =   58
               Top             =   960
               Width           =   255
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Check1"
               Height          =   270
               Index           =   0
               Left            =   -71880
               TabIndex        =   57
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check7 
               Caption         =   "Check1"
               Height          =   270
               Index           =   0
               Left            =   1560
               TabIndex        =   56
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Check1"
               Height          =   270
               Left            =   -73440
               TabIndex        =   40
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Check3"
               Height          =   270
               Index           =   5
               Left            =   -71640
               TabIndex        =   39
               Top             =   480
               Width           =   255
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Check3"
               Height          =   270
               Index           =   4
               Left            =   -71640
               TabIndex        =   38
               Top             =   1200
               Width           =   255
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Check3"
               Height          =   270
               Index           =   3
               Left            =   -71640
               TabIndex        =   37
               Top             =   1560
               Width           =   255
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Check4"
               Height          =   270
               Left            =   -73320
               TabIndex        =   36
               Top             =   480
               Width           =   225
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Cardiología"
               Height          =   375
               Index           =   43
               Left            =   -74760
               TabIndex        =   93
               Top             =   3480
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tuberculosis"
               Height          =   375
               Index           =   42
               Left            =   -74760
               TabIndex        =   92
               Top             =   3120
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Audiometría"
               Height          =   375
               Index           =   41
               Left            =   -74760
               TabIndex        =   91
               Top             =   2760
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Estadísticas"
               Height          =   375
               Index           =   37
               Left            =   -74760
               TabIndex        =   83
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresión"
               Height          =   375
               Index           =   35
               Left            =   -74880
               TabIndex        =   79
               Top             =   3840
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Familiares"
               Height          =   375
               Index           =   33
               Left            =   -74640
               TabIndex        =   70
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Exportar información"
               Height          =   375
               Index           =   32
               Left            =   -74400
               TabIndex        =   66
               Top             =   1680
               Width           =   2655
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Importar información"
               Height          =   375
               Index           =   31
               Left            =   -74040
               TabIndex        =   65
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Datos externos"
               Height          =   375
               Index           =   30
               Left            =   -74760
               TabIndex        =   64
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Editar o eliminar información"
               Height          =   375
               Index           =   29
               Left            =   -74880
               TabIndex        =   63
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Usuarios"
               Height          =   375
               Index           =   28
               Left            =   240
               TabIndex        =   55
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Usuarios"
               Height          =   375
               Index           =   25
               Left            =   -74760
               TabIndex        =   52
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Somatometría"
               Height          =   375
               Index           =   24
               Left            =   -73560
               TabIndex        =   51
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Laboratorio"
               Height          =   375
               Index           =   23
               Left            =   -73440
               TabIndex        =   50
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dental"
               Height          =   375
               Index           =   22
               Left            =   -73200
               TabIndex        =   49
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Nutrición"
               Height          =   375
               Index           =   21
               Left            =   -73200
               TabIndex        =   48
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Salud de la mujer"
               Height          =   375
               Index           =   20
               Left            =   -73800
               TabIndex        =   47
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Optometría "
               Height          =   375
               Index           =   19
               Left            =   -74760
               TabIndex        =   46
               Top             =   2400
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Editar o eliminar información"
               Height          =   375
               Index           =   18
               Left            =   -74880
               TabIndex        =   45
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Datos externos"
               Height          =   375
               Index           =   17
               Left            =   -74760
               TabIndex        =   44
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Exportar información"
               Height          =   375
               Index           =   16
               Left            =   -74400
               TabIndex        =   43
               Top             =   1560
               Width           =   2655
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Importar información"
               Height          =   375
               Index           =   15
               Left            =   -74040
               TabIndex        =   42
               Top             =   1200
               Width           =   2295
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Familiares"
               Height          =   375
               Index           =   14
               Left            =   -74640
               TabIndex        =   41
               Top             =   480
               Width           =   1215
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario"
            Height          =   375
            Index           =   27
            Left            =   240
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Contraseña"
            Height          =   375
            Index           =   26
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8655
         Begin VB.CommandButton Command1 
            Caption         =   "Guardar"
            Height          =   495
            Left            =   3840
            TabIndex        =   30
            Top             =   5880
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   390
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Height          =   390
            Index           =   0
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   2775
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   4335
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7646
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Archivo"
            TabPicture(0)   =   "Usuarios.frx":32F2
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label1(2)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Check1"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Módulos"
            TabPicture(1)   =   "Usuarios.frx":330E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Check2(9)"
            Tab(1).Control(1)=   "Check2(8)"
            Tab(1).Control(2)=   "Check2(7)"
            Tab(1).Control(3)=   "Check2(6)"
            Tab(1).Control(4)=   "Check2(5)"
            Tab(1).Control(5)=   "Check2(4)"
            Tab(1).Control(6)=   "Check2(3)"
            Tab(1).Control(7)=   "Check2(2)"
            Tab(1).Control(8)=   "Check2(1)"
            Tab(1).Control(9)=   "Check2(0)"
            Tab(1).Control(10)=   "Label1(40)"
            Tab(1).Control(11)=   "Label1(39)"
            Tab(1).Control(12)=   "Label1(38)"
            Tab(1).Control(13)=   "Label1(34)"
            Tab(1).Control(14)=   "Label1(8)"
            Tab(1).Control(15)=   "Label1(7)"
            Tab(1).Control(16)=   "Label1(6)"
            Tab(1).Control(17)=   "Label1(5)"
            Tab(1).Control(18)=   "Label1(4)"
            Tab(1).Control(19)=   "Label1(3)"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "Edición"
            TabPicture(2)   =   "Usuarios.frx":332A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label1(9)"
            Tab(2).Control(1)=   "Label1(10)"
            Tab(2).Control(2)=   "Label1(11)"
            Tab(2).Control(3)=   "Label1(12)"
            Tab(2).Control(4)=   "Check3(0)"
            Tab(2).Control(5)=   "Check3(1)"
            Tab(2).Control(6)=   "Check3(2)"
            Tab(2).ControlCount=   7
            TabCaption(3)   =   "Reportes"
            TabPicture(3)   =   "Usuarios.frx":3346
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label1(13)"
            Tab(3).Control(1)=   "Label1(36)"
            Tab(3).Control(2)=   "Check4(0)"
            Tab(3).Control(3)=   "Check4(1)"
            Tab(3).ControlCount=   4
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   9
               Left            =   -71880
               TabIndex        =   90
               Top             =   3480
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   8
               Left            =   -71880
               TabIndex        =   89
               Top             =   3120
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   7
               Left            =   -71880
               TabIndex        =   86
               Top             =   2760
               Width           =   255
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Check4"
               Height          =   270
               Index           =   1
               Left            =   -73320
               TabIndex        =   82
               Top             =   960
               Width           =   225
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   6
               Left            =   -71880
               TabIndex        =   78
               Top             =   3840
               Width           =   255
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Check4"
               Height          =   270
               Index           =   0
               Left            =   -73320
               TabIndex        =   29
               Top             =   600
               Width           =   225
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Check3"
               Height          =   270
               Index           =   2
               Left            =   -71640
               TabIndex        =   27
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Check3"
               Height          =   270
               Index           =   1
               Left            =   -71640
               TabIndex        =   26
               Top             =   1320
               Width           =   255
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Check3"
               Height          =   270
               Index           =   0
               Left            =   -71640
               TabIndex        =   25
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   5
               Left            =   -71880
               TabIndex        =   20
               Top             =   2400
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   4
               Left            =   -71880
               TabIndex        =   19
               Top             =   2040
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   3
               Left            =   -71880
               TabIndex        =   18
               Top             =   1680
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   2
               Left            =   -71880
               TabIndex        =   17
               Top             =   1320
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   1
               Left            =   -71880
               TabIndex        =   16
               Top             =   960
               Width           =   255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   270
               Index           =   0
               Left            =   -71880
               TabIndex        =   15
               Top             =   600
               Width           =   255
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Check1"
               Height          =   270
               Left            =   1560
               TabIndex        =   8
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Cardiologia"
               Height          =   375
               Index           =   40
               Left            =   -74760
               TabIndex        =   88
               Top             =   3480
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tuberculosis"
               Height          =   375
               Index           =   39
               Left            =   -74760
               TabIndex        =   87
               Top             =   3120
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Audiometría"
               Height          =   375
               Index           =   38
               Left            =   -74760
               TabIndex        =   85
               Top             =   2760
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Estadistícas"
               Height          =   375
               Index           =   36
               Left            =   -74760
               TabIndex        =   81
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Impresión"
               Height          =   375
               Index           =   34
               Left            =   -74880
               TabIndex        =   77
               Top             =   3840
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Familiares"
               Height          =   375
               Index           =   13
               Left            =   -74640
               TabIndex        =   28
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Importar información"
               Height          =   375
               Index           =   12
               Left            =   -74040
               TabIndex        =   24
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Exportar información"
               Height          =   375
               Index           =   11
               Left            =   -74400
               TabIndex        =   23
               Top             =   1680
               Width           =   2655
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Datos externos"
               Height          =   375
               Index           =   10
               Left            =   -74760
               TabIndex        =   22
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Editar o eliminar información"
               Height          =   375
               Index           =   9
               Left            =   -74880
               TabIndex        =   21
               Top             =   600
               Width           =   3135
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Optometría"
               Height          =   375
               Index           =   8
               Left            =   -74760
               TabIndex        =   14
               Top             =   2400
               Width           =   2775
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Salud de la mujer"
               Height          =   375
               Index           =   7
               Left            =   -73800
               TabIndex        =   13
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Nutrición"
               Height          =   375
               Index           =   6
               Left            =   -73200
               TabIndex        =   12
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dental"
               Height          =   375
               Index           =   5
               Left            =   -73200
               TabIndex        =   11
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Laboratorio"
               Height          =   375
               Index           =   4
               Left            =   -73440
               TabIndex        =   10
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Somatometría"
               Height          =   375
               Index           =   3
               Left            =   -73560
               TabIndex        =   9
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Usuarios"
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   7
               Top             =   600
               Width           =   1215
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Contraseña"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    If Not Usuarios.Text1(0) = "" And Not Usuarios.Text1(1) = "" Then
        With RsUsers
            .Requery
            .AddNew
                .Fields("NOMBRE") = Usuarios.Text1(0).Text
                .Fields("PASS") = Usuarios.Text1(1).Text
                .Fields("USUARIOS") = Usuarios.Check1.Value
                .Fields("SOMATOMETRIA") = Usuarios.Check2(0).Value
                .Fields("LABORATORIO") = Usuarios.Check2(1).Value
                .Fields("DENTAL") = Usuarios.Check2(2).Value
                .Fields("NUTRICION") = Usuarios.Check2(3).Value
                .Fields("MUJER") = Usuarios.Check2(4).Value
                .Fields("OPTOMETRIA") = Usuarios.Check2(5).Value
                .Fields("IMPRESION") = Usuarios.Check2(6).Value
                .Fields("AUDIOMETRIA") = Usuarios.Check2(7).Value
                .Fields("TUBERCULOSIS") = Usuarios.Check2(8).Value
                .Fields("CARDIOLOGIA") = Usuarios.Check2(9).Value
                .Fields("EDITAR") = Usuarios.Check3(0).Value
                .Fields("IMPORTAR") = Usuarios.Check3(1).Value
                .Fields("EXPORTAR") = Usuarios.Check3(2).Value
                .Fields("FAMILIARES") = Usuarios.Check4(0).Value
                .Fields("ESTADISTICAS") = Usuarios.Check4(1).Value
            .Update
            .Requery
        End With
        MsgBox ("Ususario " & Usuarios.Text1(0) & " creado con éxito"), vbOKOnly, "Usuario"
        Usuarios.Text1(0).Text = ""
        Usuarios.Text1(1).Text = ""
        Usuarios.Check1.Value = 0
        Usuarios.Check2(0).Value = 0
        Usuarios.Check2(1).Value = 0
        Usuarios.Check2(2).Value = 0
        Usuarios.Check2(3).Value = 0
        Usuarios.Check2(4).Value = 0
        Usuarios.Check2(5).Value = 0
        Usuarios.Check2(6).Value = 0
        Usuarios.Check2(7).Value = 0
        Usuarios.Check2(8).Value = 0
        Usuarios.Check2(9).Value = 0
        Usuarios.Check3(0).Value = 0
        Usuarios.Check3(1).Value = 0
        Usuarios.Check3(2).Value = 0
        Usuarios.Check4(0).Value = 0
        Usuarios.Check4(1).Value = 0
    Else
        MsgBox ("Usuario y contraseña necesarios para crear el usuario"), vbOKOnly, "Error"
    End If
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    With RsUsers
        .Update
        MsgBox ("Usuario actualizado"), vbOKOnly, "Completado"
    End With
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error Resume Next
    With RsUsers
        Select Case Index
            Case 0
                .MoveFirst
            Case 1
                .MovePrevious
            Case 2
                .MoveNext
            Case 3
                .MoveLast
            Case 4
                .Delete
                MsgBox ("Usuario eliminado"), vbOKOnly, "Completado"
        End Select
    End With
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RsUsers
        If .State = 1 Then .Close
            .Open "Select * from USERS", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Set Usuarios.Text1(2).DataSource = RsUsers
    Set Usuarios.Text1(3).DataSource = RsUsers
    Set Usuarios.Check7(0).DataSource = RsUsers
    Set Usuarios.Check8(0).DataSource = RsUsers
    Set Usuarios.Check8(1).DataSource = RsUsers
    Set Usuarios.Check8(2).DataSource = RsUsers
    Set Usuarios.Check8(3).DataSource = RsUsers
    Set Usuarios.Check8(4).DataSource = RsUsers
    Set Usuarios.Check8(5).DataSource = RsUsers
    Set Usuarios.Check8(6).DataSource = RsUsers
    Set Usuarios.Check8(7).DataSource = RsUsers
    Set Usuarios.Check8(8).DataSource = RsUsers
    Set Usuarios.Check8(9).DataSource = RsUsers
    Set Usuarios.Check9(0).DataSource = RsUsers
    Set Usuarios.Check9(1).DataSource = RsUsers
    Set Usuarios.Check9(2).DataSource = RsUsers
    Set Usuarios.Check10(0).DataSource = RsUsers
    Set Usuarios.Check10(1).DataSource = RsUsers
    Usuarios.Text1(2).DataField = ("NOMBRE")
    Usuarios.Text1(3).DataField = ("PASS")
    Usuarios.Check7(0).DataField = ("USUARIOS")
    Usuarios.Check8(0).DataField = ("SOMATOMETRIA")
    Usuarios.Check8(1).DataField = ("LABORATORIO")
    Usuarios.Check8(2).DataField = ("DENTAL")
    Usuarios.Check8(3).DataField = ("NUTRICION")
    Usuarios.Check8(4).DataField = ("MUJER")
    Usuarios.Check8(5).DataField = ("OPTOMETRIA")
    Usuarios.Check8(6).DataField = ("IMPRESION")
    Usuarios.Check8(7).DataField = ("AUDIOMETRIA")
    Usuarios.Check8(8).DataField = ("TUBERCULOSIS")
    Usuarios.Check8(9).DataField = ("CARDIOLOGIA")
    Usuarios.Check9(0).DataField = ("EDITAR")
    Usuarios.Check9(1).DataField = ("IMPORTAR")
    Usuarios.Check9(2).DataField = ("EXPORTAR")
    Usuarios.Check10(0).DataField = ("FAMILIARES")
    Usuarios.Check10(1).DataField = ("ESTADISTICAS")
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    Usuarios.Text1(0).Text = ""
    Usuarios.Text1(1).Text = ""
    Usuarios.Check1.Value = 0
    Usuarios.Check2(0).Value = 0
    Usuarios.Check2(1).Value = 0
    Usuarios.Check2(2).Value = 0
    Usuarios.Check2(3).Value = 0
    Usuarios.Check2(4).Value = 0
    Usuarios.Check2(5).Value = 0
    Usuarios.Check2(6).Value = 0
    Usuarios.Check2(7).Value = 0
    Usuarios.Check2(8).Value = 0
    Usuarios.Check2(9).Value = 0
    Usuarios.Check3(0).Value = 0
    Usuarios.Check3(1).Value = 0
    Usuarios.Check3(2).Value = 0
    Usuarios.Check4(0).Value = 0
    Usuarios.Check4(1).Value = 0
    Unload Me
    Form1.Enabled = True
End Sub

