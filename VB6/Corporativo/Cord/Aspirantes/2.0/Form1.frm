VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicios Cord SA de CV - Solicitud de Empleo"
   ClientHeight    =   10440
   ClientLeft      =   -5265
   ClientTop       =   -2055
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar formulario"
      Height          =   390
      Index           =   1
      Left            =   10560
      TabIndex        =   266
      Top             =   9840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Terminado"
      Height          =   390
      Index           =   0
      Left            =   120
      TabIndex        =   265
      Top             =   9840
      Width           =   2775
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   1058
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal"
      TabPicture(0)   =   "Form1.frx":10CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(15)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(17)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(18)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(19)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(20)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Image1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Combo1(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo1(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "DTPicker1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Combo1(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(4)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(6)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(7)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(8)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Combo1(3)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Combo1(4)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(9)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(10)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(11)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Combo1(5)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(12)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(13)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Command2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Familiar"
      TabPicture(1)   =   "Form1.frx":10E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(0)"
      Tab(1).Control(1)=   "Label2(1)"
      Tab(1).Control(2)=   "Label2(2)"
      Tab(1).Control(3)=   "Label2(3)"
      Tab(1).Control(4)=   "Label2(4)"
      Tab(1).Control(5)=   "DTPicker2(18)"
      Tab(1).Control(6)=   "DTPicker2(17)"
      Tab(1).Control(7)=   "DTPicker2(16)"
      Tab(1).Control(8)=   "DTPicker2(15)"
      Tab(1).Control(9)=   "DTPicker2(14)"
      Tab(1).Control(10)=   "DTPicker2(13)"
      Tab(1).Control(11)=   "DTPicker2(12)"
      Tab(1).Control(12)=   "DTPicker2(11)"
      Tab(1).Control(13)=   "DTPicker2(10)"
      Tab(1).Control(14)=   "DTPicker2(9)"
      Tab(1).Control(15)=   "DTPicker2(8)"
      Tab(1).Control(16)=   "DTPicker2(7)"
      Tab(1).Control(17)=   "DTPicker2(6)"
      Tab(1).Control(18)=   "DTPicker2(5)"
      Tab(1).Control(19)=   "DTPicker2(4)"
      Tab(1).Control(20)=   "DTPicker2(3)"
      Tab(1).Control(21)=   "DTPicker2(2)"
      Tab(1).Control(22)=   "DTPicker2(1)"
      Tab(1).Control(23)=   "Text2(0)"
      Tab(1).Control(24)=   "Combo2(0)"
      Tab(1).Control(25)=   "DTPicker2(0)"
      Tab(1).Control(26)=   "Text2(1)"
      Tab(1).Control(27)=   "Text2(2)"
      Tab(1).Control(28)=   "Text2(3)"
      Tab(1).Control(29)=   "Combo2(1)"
      Tab(1).Control(30)=   "Text2(4)"
      Tab(1).Control(31)=   "Text2(5)"
      Tab(1).Control(32)=   "Text2(6)"
      Tab(1).Control(33)=   "Combo2(2)"
      Tab(1).Control(34)=   "Text2(7)"
      Tab(1).Control(35)=   "Text2(8)"
      Tab(1).Control(36)=   "Text2(9)"
      Tab(1).Control(37)=   "Combo2(3)"
      Tab(1).Control(38)=   "Text2(10)"
      Tab(1).Control(39)=   "Text2(11)"
      Tab(1).Control(40)=   "Text2(12)"
      Tab(1).Control(41)=   "Combo2(4)"
      Tab(1).Control(42)=   "Text2(13)"
      Tab(1).Control(43)=   "Text2(14)"
      Tab(1).Control(44)=   "Text2(15)"
      Tab(1).Control(45)=   "Combo2(5)"
      Tab(1).Control(46)=   "Text2(16)"
      Tab(1).Control(47)=   "Text2(17)"
      Tab(1).Control(48)=   "Text2(18)"
      Tab(1).Control(49)=   "Combo2(6)"
      Tab(1).Control(50)=   "Text2(19)"
      Tab(1).Control(51)=   "Text2(20)"
      Tab(1).Control(52)=   "Text2(21)"
      Tab(1).Control(53)=   "Combo2(7)"
      Tab(1).Control(54)=   "Text2(22)"
      Tab(1).Control(55)=   "Text2(23)"
      Tab(1).Control(56)=   "Text2(24)"
      Tab(1).Control(57)=   "Combo2(8)"
      Tab(1).Control(58)=   "Text2(25)"
      Tab(1).Control(59)=   "Text2(26)"
      Tab(1).Control(60)=   "Text2(27)"
      Tab(1).Control(61)=   "Combo2(9)"
      Tab(1).Control(62)=   "Text2(28)"
      Tab(1).Control(63)=   "Text2(29)"
      Tab(1).Control(64)=   "Text2(30)"
      Tab(1).Control(65)=   "Combo2(10)"
      Tab(1).Control(66)=   "Text2(31)"
      Tab(1).Control(67)=   "Text2(32)"
      Tab(1).Control(68)=   "Text2(33)"
      Tab(1).Control(69)=   "Combo2(12)"
      Tab(1).Control(70)=   "Text2(34)"
      Tab(1).Control(71)=   "Text2(35)"
      Tab(1).Control(72)=   "Text2(36)"
      Tab(1).Control(73)=   "Combo2(11)"
      Tab(1).Control(74)=   "Text2(37)"
      Tab(1).Control(75)=   "Text2(38)"
      Tab(1).Control(76)=   "Text2(39)"
      Tab(1).Control(77)=   "Combo2(13)"
      Tab(1).Control(78)=   "Text2(40)"
      Tab(1).Control(79)=   "Text2(41)"
      Tab(1).Control(80)=   "Text2(42)"
      Tab(1).Control(81)=   "Combo2(14)"
      Tab(1).Control(82)=   "Text2(43)"
      Tab(1).Control(83)=   "Text2(44)"
      Tab(1).Control(84)=   "Text2(45)"
      Tab(1).Control(85)=   "Combo2(15)"
      Tab(1).Control(86)=   "Text2(46)"
      Tab(1).Control(87)=   "Text2(47)"
      Tab(1).Control(88)=   "Text2(48)"
      Tab(1).Control(89)=   "Combo2(16)"
      Tab(1).Control(90)=   "Text2(49)"
      Tab(1).Control(91)=   "Text2(50)"
      Tab(1).Control(92)=   "Text2(51)"
      Tab(1).Control(93)=   "Combo2(17)"
      Tab(1).Control(94)=   "Text2(52)"
      Tab(1).Control(95)=   "Text2(53)"
      Tab(1).Control(96)=   "Text2(54)"
      Tab(1).Control(97)=   "Combo2(18)"
      Tab(1).Control(98)=   "Text2(55)"
      Tab(1).Control(99)=   "Text2(56)"
      Tab(1).ControlCount=   100
      TabCaption(2)   =   "Estado de Salud"
      TabPicture(2)   =   "Form1.frx":1102
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check1(3)"
      Tab(2).Control(1)=   "Text3(5)"
      Tab(2).Control(2)=   "Check1(2)"
      Tab(2).Control(3)=   "Text3(4)"
      Tab(2).Control(4)=   "Text3(3)"
      Tab(2).Control(5)=   "Text3(2)"
      Tab(2).Control(6)=   "Text3(1)"
      Tab(2).Control(7)=   "Check1(1)"
      Tab(2).Control(8)=   "Text3(0)"
      Tab(2).Control(9)=   "Check1(0)"
      Tab(2).Control(10)=   "Label3(9)"
      Tab(2).Control(11)=   "Label3(8)"
      Tab(2).Control(12)=   "Label3(7)"
      Tab(2).Control(13)=   "Label3(6)"
      Tab(2).Control(14)=   "Label3(5)"
      Tab(2).Control(15)=   "Label3(4)"
      Tab(2).Control(16)=   "Label3(3)"
      Tab(2).Control(17)=   "Label3(2)"
      Tab(2).Control(18)=   "Label3(1)"
      Tab(2).Control(19)=   "Label3(0)"
      Tab(2).ControlCount=   20
      TabCaption(3)   =   "Conocimientos Generales"
      TabPicture(3)   =   "Form1.frx":111E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label4(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label4(2)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label4(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label4(4)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Text4(0)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Text4(1)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Check2(0)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Check2(1)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Check2(2)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Text4(2)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Empleos Anteriores"
      TabPicture(4)   =   "Form1.frx":113A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5(0)"
      Tab(4).Control(1)=   "Label5(1)"
      Tab(4).Control(2)=   "Label5(2)"
      Tab(4).Control(3)=   "Label5(3)"
      Tab(4).Control(4)=   "Frame1(0)"
      Tab(4).Control(5)=   "Check3"
      Tab(4).Control(6)=   "Text5(0)"
      Tab(4).Control(7)=   "Text5(1)"
      Tab(4).Control(8)=   "Frame1(1)"
      Tab(4).Control(9)=   "Text5(2)"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Puesto que Solicita"
      TabPicture(5)   =   "Form1.frx":1156
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label6(0)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label6(1)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label6(2)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label6(3)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label6(4)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Text6(0)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Text6(1)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Text6(2)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Text6(3)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Text6(4)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "General"
      TabPicture(6)   =   "Form1.frx":1172
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Check4(2)"
      Tab(6).Control(1)=   "Check4(1)"
      Tab(6).Control(2)=   "Check4(0)"
      Tab(6).Control(3)=   "Text7"
      Tab(6).Control(4)=   "Label7(3)"
      Tab(6).Control(5)=   "Label7(2)"
      Tab(6).Control(6)=   "Label7(1)"
      Tab(6).Control(7)=   "Label7(0)"
      Tab(6).ControlCount=   8
      TabCaption(7)   =   "Económico"
      TabPicture(7)   =   "Form1.frx":118E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Check5(3)"
      Tab(7).Control(1)=   "Check5(2)"
      Tab(7).Control(2)=   "Check5(1)"
      Tab(7).Control(3)=   "Check5(0)"
      Tab(7).Control(4)=   "Label8(3)"
      Tab(7).Control(5)=   "Label8(2)"
      Tab(7).Control(6)=   "Label8(1)"
      Tab(7).Control(7)=   "Label8(0)"
      Tab(7).ControlCount=   8
      TabCaption(8)   =   "Tiempo libre"
      TabPicture(8)   =   "Form1.frx":11AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Text8(2)"
      Tab(8).Control(1)=   "Text8(1)"
      Tab(8).Control(2)=   "Check6(1)"
      Tab(8).Control(3)=   "Check6(0)"
      Tab(8).Control(4)=   "Text8(0)"
      Tab(8).Control(5)=   "Label9(4)"
      Tab(8).Control(6)=   "Label9(3)"
      Tab(8).Control(7)=   "Label9(2)"
      Tab(8).Control(8)=   "Label9(1)"
      Tab(8).Control(9)=   "Label9(0)"
      Tab(8).ControlCount=   10
      TabCaption(9)   =   "Dirección"
      TabPicture(9)   =   "Form1.frx":11C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Text9(6)"
      Tab(9).Control(1)=   "Combo3"
      Tab(9).Control(2)=   "Text9(5)"
      Tab(9).Control(3)=   "Text9(4)"
      Tab(9).Control(4)=   "Text9(3)"
      Tab(9).Control(5)=   "Text9(2)"
      Tab(9).Control(6)=   "Text9(1)"
      Tab(9).Control(7)=   "Text9(0)"
      Tab(9).Control(8)=   "Label1(28)"
      Tab(9).Control(9)=   "Label1(27)"
      Tab(9).Control(10)=   "Label1(26)"
      Tab(9).Control(11)=   "Label1(25)"
      Tab(9).Control(12)=   "Label1(24)"
      Tab(9).Control(13)=   "Label1(23)"
      Tab(9).Control(14)=   "Label1(22)"
      Tab(9).Control(15)=   "Label1(21)"
      Tab(9).ControlCount=   16
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -70080
         TabIndex        =   173
         Top             =   3360
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   56
         Left            =   -64680
         TabIndex        =   137
         Top             =   8400
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   55
         Left            =   -67200
         TabIndex        =   136
         Top             =   8400
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   18
         Left            =   -71280
         TabIndex        =   134
         Top             =   8400
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   54
         Left            =   -74640
         TabIndex        =   133
         Top             =   8400
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   53
         Left            =   -64680
         TabIndex        =   132
         Top             =   8040
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   52
         Left            =   -67200
         TabIndex        =   131
         Top             =   8040
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   17
         Left            =   -71280
         TabIndex        =   129
         Top             =   8040
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   51
         Left            =   -74640
         TabIndex        =   128
         Top             =   8040
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   50
         Left            =   -64680
         TabIndex        =   127
         Top             =   7680
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   49
         Left            =   -67200
         TabIndex        =   126
         Top             =   7680
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   16
         Left            =   -71280
         TabIndex        =   124
         Top             =   7680
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   48
         Left            =   -74640
         TabIndex        =   123
         Top             =   7680
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   47
         Left            =   -64680
         TabIndex        =   122
         Top             =   7320
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   46
         Left            =   -67200
         TabIndex        =   121
         Top             =   7320
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   15
         Left            =   -71280
         TabIndex        =   119
         Top             =   7320
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   45
         Left            =   -74640
         MaxLength       =   10
         TabIndex        =   118
         Top             =   7320
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   44
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   117
         Top             =   6960
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   43
         Left            =   -67200
         TabIndex        =   116
         Top             =   6960
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   14
         Left            =   -71280
         TabIndex        =   114
         Top             =   6960
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   42
         Left            =   -74640
         TabIndex        =   113
         Top             =   6960
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Tomar foto"
         Height          =   1110
         Left            =   7680
         TabIndex        =   42
         Top             =   8040
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   6
         Left            =   -66480
         TabIndex        =   264
         Top             =   3960
         Width           =   4215
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         Height          =   390
         ItemData        =   "Form1.frx":11E2
         Left            =   -72840
         List            =   "Form1.frx":11E4
         TabIndex        =   263
         Top             =   3960
         Width           =   4335
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   5
         Left            =   -72840
         TabIndex        =   262
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   -72840
         TabIndex        =   261
         Top             =   3000
         Width           =   10575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   -72840
         TabIndex        =   260
         Top             =   2520
         Width           =   10575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -66480
         TabIndex        =   259
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -72840
         TabIndex        =   258
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -72840
         TabIndex        =   257
         Top             =   1560
         Width           =   10575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   13
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   34
         Top             =   4440
         Width           =   3855
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -68160
         TabIndex        =   247
         Top             =   3360
         Width           =   6135
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -68160
         TabIndex        =   246
         Top             =   2880
         Width           =   6135
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -68160
         TabIndex        =   245
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -68160
         TabIndex        =   244
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -68160
         TabIndex        =   243
         Top             =   1440
         Width           =   6135
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   -71400
         TabIndex        =   237
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   -71400
         TabIndex        =   236
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -71400
         TabIndex        =   235
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -71400
         TabIndex        =   234
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   -69600
         TabIndex        =   229
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -69600
         TabIndex        =   228
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -69600
         TabIndex        =   227
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   -69600
         TabIndex        =   226
         Top             =   1440
         Width           =   7575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   -67920
         TabIndex        =   221
         Top             =   3360
         Width           =   5895
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   -67920
         TabIndex        =   220
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -67920
         TabIndex        =   219
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -67920
         TabIndex        =   218
         Top             =   1920
         Width           =   5895
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -67920
         TabIndex        =   217
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -69960
         TabIndex        =   189
         Top             =   2880
         Width           =   7815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Empleo anterior"
         Height          =   3975
         Index           =   1
         Left            =   -68280
         TabIndex        =   197
         Top             =   3360
         Width           =   6135
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   16
            Left            =   1800
            TabIndex        =   204
            Top             =   3360
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   15
            Left            =   1800
            TabIndex        =   203
            Top             =   2640
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   14
            Left            =   1800
            TabIndex        =   202
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   13
            Left            =   1800
            TabIndex        =   201
            Top             =   1680
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   12
            Left            =   1800
            TabIndex        =   200
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   11
            Left            =   1800
            TabIndex        =   199
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   10
            Left            =   1800
            TabIndex        =   198
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   211
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Domicilio:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   210
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Tiempo:"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   209
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Encargado:"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   208
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Actividades:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   207
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Sueldo:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   206
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Motivo de renuncia:"
            Height          =   615
            Index           =   11
            Left            =   120
            TabIndex        =   205
            Top             =   3240
            Width           =   1575
         End
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -69960
         TabIndex        =   188
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -69960
         TabIndex        =   187
         Top             =   1920
         Width           =   7815
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -69960
         TabIndex        =   186
         Top             =   1560
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Empleo actual"
         Height          =   3975
         Index           =   0
         Left            =   -74640
         TabIndex        =   178
         Top             =   3360
         Width           =   6135
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   9
            Left            =   1800
            TabIndex        =   196
            Top             =   3360
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   8
            Left            =   1800
            TabIndex        =   195
            Top             =   2640
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   7
            Left            =   1800
            TabIndex        =   194
            Top             =   2160
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   6
            Left            =   1800
            TabIndex        =   193
            Top             =   1680
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   5
            Left            =   1800
            TabIndex        =   192
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   4
            Left            =   1800
            TabIndex        =   191
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Index           =   3
            Left            =   1800
            TabIndex        =   190
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Motivo de renuncia:"
            Height          =   615
            Index           =   10
            Left            =   120
            TabIndex        =   185
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Sueldo:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   184
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Actividades:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   183
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Encargado:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   182
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Tiempo:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   181
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Domicilio:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   180
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Empresa:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   179
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   -70560
         TabIndex        =   172
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -70560
         TabIndex        =   171
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -70560
         TabIndex        =   170
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -70560
         TabIndex        =   169
         Top             =   1920
         Width           =   8415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -70560
         TabIndex        =   168
         Top             =   1440
         Width           =   8415
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   -70440
         TabIndex        =   162
         Top             =   6120
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   -70440
         TabIndex        =   161
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   -70440
         TabIndex        =   160
         Top             =   5160
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   -70440
         TabIndex        =   159
         Top             =   4560
         Width           =   8415
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   -70440
         TabIndex        =   158
         Top             =   4080
         Width           =   8415
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -70440
         TabIndex        =   157
         Top             =   3360
         Width           =   8415
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -70440
         TabIndex        =   156
         Top             =   2880
         Width           =   8415
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -70440
         TabIndex        =   155
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -70440
         TabIndex        =   146
         Top             =   1920
         Width           =   8415
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -70440
         TabIndex        =   145
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   41
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   112
         Top             =   6600
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   40
         Left            =   -67200
         TabIndex        =   111
         Top             =   6600
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   13
         Left            =   -71280
         TabIndex        =   109
         Top             =   6600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   39
         Left            =   -74640
         TabIndex        =   108
         Top             =   6600
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   38
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   107
         Top             =   6240
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   37
         Left            =   -67200
         TabIndex        =   106
         Top             =   6240
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   11
         Left            =   -71280
         TabIndex        =   104
         Top             =   6240
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   36
         Left            =   -74640
         TabIndex        =   103
         Top             =   6240
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   35
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   102
         Top             =   5880
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   34
         Left            =   -67200
         TabIndex        =   101
         Top             =   5880
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   12
         Left            =   -71280
         TabIndex        =   99
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   33
         Left            =   -74640
         TabIndex        =   98
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   32
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   97
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   31
         Left            =   -67200
         TabIndex        =   96
         Top             =   5520
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   10
         Left            =   -71280
         TabIndex        =   94
         Top             =   5520
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   30
         Left            =   -74640
         TabIndex        =   93
         Top             =   5520
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   29
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   92
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   28
         Left            =   -67200
         TabIndex        =   91
         Top             =   5160
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   9
         Left            =   -71280
         TabIndex        =   89
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   27
         Left            =   -74640
         TabIndex        =   88
         Top             =   5160
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   26
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   87
         Top             =   4800
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   25
         Left            =   -67200
         TabIndex        =   86
         Top             =   4800
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   8
         Left            =   -71280
         TabIndex        =   84
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   24
         Left            =   -74640
         TabIndex        =   83
         Top             =   4800
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   23
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   82
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   22
         Left            =   -67200
         TabIndex        =   81
         Top             =   4440
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   7
         Left            =   -71280
         TabIndex        =   79
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   21
         Left            =   -74640
         TabIndex        =   78
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   20
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   77
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   19
         Left            =   -67200
         TabIndex        =   76
         Top             =   4080
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   6
         Left            =   -71280
         TabIndex        =   74
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   18
         Left            =   -74640
         TabIndex        =   73
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   17
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   72
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   16
         Left            =   -67200
         TabIndex        =   71
         Top             =   3720
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   5
         Left            =   -71280
         TabIndex        =   69
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   15
         Left            =   -74640
         TabIndex        =   68
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   14
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   67
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   13
         Left            =   -67200
         TabIndex        =   66
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   -71280
         TabIndex        =   64
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   12
         Left            =   -74640
         TabIndex        =   63
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   11
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   62
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   10
         Left            =   -67200
         TabIndex        =   61
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   -71280
         TabIndex        =   59
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   9
         Left            =   -74640
         TabIndex        =   58
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   8
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   57
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   7
         Left            =   -67200
         TabIndex        =   56
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -71280
         TabIndex        =   54
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   6
         Left            =   -74640
         TabIndex        =   53
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   5
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   52
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   -67200
         TabIndex        =   51
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -71280
         TabIndex        =   49
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   -74640
         TabIndex        =   48
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   -64680
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   -67200
         TabIndex        =   46
         Top             =   1920
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   0
         Left            =   -68880
         TabIndex        =   45
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -71280
         TabIndex        =   44
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   -74640
         TabIndex        =   43
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   390
         Index           =   12
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   41
         Top             =   6840
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   5
         Left            =   9000
         TabIndex        =   40
         Top             =   6360
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   39
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   10
         Left            =   9000
         TabIndex        =   38
         Top             =   5880
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   9
         Left            =   2640
         TabIndex        =   37
         Top             =   5880
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   2640
         TabIndex        =   36
         Top             =   5400
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   2640
         TabIndex        =   35
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   8
         Left            =   2640
         TabIndex        =   33
         Top             =   4440
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   7
         Left            =   9000
         TabIndex        =   32
         Top             =   3960
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   6
         Left            =   2640
         TabIndex        =   31
         Top             =   3960
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   5
         Left            =   9000
         TabIndex        =   30
         Top             =   3480
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   4
         Left            =   2640
         TabIndex        =   29
         Top             =   3480
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   9000
         TabIndex        =   28
         Top             =   3000
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2640
         TabIndex        =   27
         Top             =   3000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   9000
         TabIndex        =   26
         Top             =   2520
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         ItemData        =   "Form1.frx":11E6
         Left            =   2640
         List            =   "Form1.frx":11E8
         TabIndex        =   25
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   3
         Left            =   9000
         TabIndex        =   9
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   2
         Left            =   2640
         TabIndex        =   8
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   1
         Left            =   9000
         TabIndex        =   7
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   390
         Index           =   0
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   1
         Left            =   -68880
         TabIndex        =   50
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   2
         Left            =   -68880
         TabIndex        =   55
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   3
         Left            =   -68880
         TabIndex        =   60
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   4
         Left            =   -68880
         TabIndex        =   65
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   5
         Left            =   -68880
         TabIndex        =   70
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   6
         Left            =   -68880
         TabIndex        =   75
         Top             =   4080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   7
         Left            =   -68880
         TabIndex        =   80
         Top             =   4440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   8
         Left            =   -68880
         TabIndex        =   85
         Top             =   4800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   9
         Left            =   -68880
         TabIndex        =   90
         Top             =   5160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   10
         Left            =   -68880
         TabIndex        =   95
         Top             =   5520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   11
         Left            =   -68880
         TabIndex        =   100
         Top             =   5880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   12
         Left            =   -68880
         TabIndex        =   105
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   13
         Left            =   -68880
         TabIndex        =   110
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   14
         Left            =   -68880
         TabIndex        =   115
         Top             =   6960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   15
         Left            =   -68880
         TabIndex        =   120
         Top             =   7320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   16
         Left            =   -68880
         TabIndex        =   125
         Top             =   7680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   17
         Left            =   -68880
         TabIndex        =   130
         Top             =   8040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Index           =   18
         Left            =   -68880
         TabIndex        =   135
         Top             =   8400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119013377
         CurrentDate     =   43228
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   10560
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pais"
         Height          =   255
         Index           =   28
         Left            =   -68520
         TabIndex        =   256
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   255
         Index           =   27
         Left            =   -74880
         TabIndex        =   255
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Postal"
         Height          =   255
         Index           =   26
         Left            =   -74880
         TabIndex        =   254
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Index           =   25
         Left            =   -74880
         TabIndex        =   253
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Número interior"
         Height          =   255
         Index           =   24
         Left            =   -68520
         TabIndex        =   252
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Colonia"
         Height          =   255
         Index           =   23
         Left            =   -74880
         TabIndex        =   251
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Exterior"
         Height          =   255
         Index           =   22
         Left            =   -74880
         TabIndex        =   250
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Calle"
         Height          =   255
         Index           =   21
         Left            =   -74880
         TabIndex        =   249
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         Height          =   255
         Index           =   20
         Left            =   6600
         TabIndex        =   248
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mencione al menos 2 cosas negativas sobre usted:"
         Height          =   375
         Index           =   4
         Left            =   -74640
         TabIndex        =   242
         Top             =   3480
         Width           =   6375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mencione al menos 2 cosas positivas sobre usted:"
         Height          =   375
         Index           =   3
         Left            =   -73560
         TabIndex        =   241
         Top             =   3000
         Width           =   5295
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Usted se considera una persona con valores y principios?"
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   240
         Top             =   2520
         Width           =   6375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Practica algún deporte?"
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   239
         Top             =   2040
         Width           =   6375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mencione la actividad que mas gusta de realizar:"
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   238
         Top             =   1560
         Width           =   6375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Tiene deudas?"
         Height          =   375
         Index           =   3
         Left            =   -74760
         TabIndex        =   233
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Su pareja trabaja?"
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   232
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Posee casa propia?"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   231
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Tiene usted otros ingresos?"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   230
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Tiene mascotas?"
         Height          =   375
         Index           =   3
         Left            =   -74760
         TabIndex        =   225
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Posee algun vehículo?"
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   224
         Top             =   2520
         Width           =   4935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Tiene familiares laborando en esta empresa?"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   223
         Top             =   2040
         Width           =   4935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Cómo se enteró de esta empresa?"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   222
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Con qué tipo de compañeros le gustaría laborar?"
         Height          =   375
         Index           =   4
         Left            =   -74760
         TabIndex        =   216
         Top             =   3480
         Width           =   6615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Cuánto le gustaría ganar?"
         Height          =   375
         Index           =   3
         Left            =   -74760
         TabIndex        =   215
         Top             =   3000
         Width           =   6615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Qué habilidades posee para desempeñarse en ese puesto?"
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   214
         Top             =   2520
         Width           =   6615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Porque cree usted que es apto para el puesto?"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   213
         Top             =   2040
         Width           =   6615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¿Qué puesto solicita?"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   212
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo de su renuncia:"
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   177
         Top             =   3000
         Width           =   4455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "¿Durante cuánto tiempo?"
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   176
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Área en la que laboró:"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   175
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "¿Laboró anteriormente en esta empresa?"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   174
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sabe conducir:"
         Height          =   255
         Index           =   4
         Left            =   -73920
         TabIndex        =   167
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sabe escribir:"
         Height          =   255
         Index           =   3
         Left            =   -73920
         TabIndex        =   166
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sabe leer:"
         Height          =   255
         Index           =   2
         Left            =   -73920
         TabIndex        =   165
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Programas informáticos que domina:"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   164
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Idiomas que domina:"
         Height          =   255
         Index           =   0
         Left            =   -73920
         TabIndex        =   163
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "¿Toma alcohol con frecuencia?"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   -74640
         TabIndex        =   154
         Top             =   6120
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de cigarrillos diarios:"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   -74640
         TabIndex        =   153
         Top             =   5640
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "¿Fuma?"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   -74640
         TabIndex        =   152
         Top             =   5160
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alergias:"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   -74640
         TabIndex        =   151
         Top             =   4680
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "¿Actualmente toma algún medicamento?"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   -74640
         TabIndex        =   150
         Top             =   3960
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enfermedades crónicas:"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   -74640
         TabIndex        =   149
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Explique el motivo brevemente:"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   -74640
         TabIndex        =   148
         Top             =   3000
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "¿Alguna vez estuvo internado?"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   -74640
         TabIndex        =   147
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Explíquela brevemente:"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   -74640
         TabIndex        =   144
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "¿Ha tenido alguna lesión y/o fractura?"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   143
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teléfono"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   -64680
         TabIndex        =   142
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Domicilio"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   -67200
         TabIndex        =   141
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F. Nacimiento"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   -68880
         TabIndex        =   140
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parentesco"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   -71280
         TabIndex        =   139
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   138
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Talla de Zapatos:"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   24
         Top             =   6960
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Talla de Camisa:"
         Height          =   255
         Index           =   18
         Left            =   6600
         TabIndex        =   23
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Talla de Pantalón:"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   22
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula Profesional:"
         Height          =   255
         Index           =   16
         Left            =   6600
         TabIndex        =   21
         Top             =   6000
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Especialidad:"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   20
         Top             =   6000
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Escolaridad:"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   19
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Sangre:"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   18
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Correo:"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   17
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "IMSS:"
         Height          =   255
         Index           =   11
         Left            =   6600
         TabIndex        =   16
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "IFE:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   15
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CURP:"
         Height          =   255
         Index           =   9
         Left            =   6600
         TabIndex        =   14
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "RFC:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
         Height          =   255
         Index           =   7
         Left            =   6600
         TabIndex        =   12
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Nacimiento:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad:"
         Height          =   255
         Index           =   5
         Left            =   6600
         TabIndex        =   10
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Materno:"
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido Paterno:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Nombre:"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Nombre:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   1680
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Función SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Check1(0).Value = 0 Then
                Text3(0) = ""
                Text3(0).Enabled = False
            Else
                Text3(0).Enabled = True
            End If
        Case 1
            If Check1(1).Value = 0 Then
                Text3(1) = ""
                Text3(1).Enabled = False
            Else
                Text3(1).Enabled = True
            End If
        Case 2
            If Check1(2).Value = 0 Then
                Text3(5) = ""
                Text3(5).Enabled = False
            Else
                Text3(5).Enabled = True
            End If
    End Select
End Sub

Private Sub Check2_Click(Index As Integer)
Select Case Index
        Case 2
            If Check2(2).Value = 0 Then
                Text4(2) = ""
                Text4(2).Enabled = False
            Else
                Text4(2).Enabled = True
            End If
    End Select
End Sub

Private Sub Check3_Click()
    If Check3.Value = 0 Then
        Text5(0) = ""
        Text5(1) = ""
        Text5(2) = ""
        Text5(0).Enabled = False
        Text5(1).Enabled = False
        Text5(2).Enabled = False
    Else
        Text5(0).Enabled = True
        Text5(1).Enabled = True
        Text5(2).Enabled = True
    End If
End Sub

Private Sub Combo1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Combo1(0).Text = "CASADO" Or Combo1(0).Text = "UNION LIBRE" Then
                Check5(2).Enabled = True
            Else
                Check5(2).Value = 0
                Check5(2).Enabled = False
            End If
        Case 4
            If Combo1(4).Text = "DOCTORADO" Or Combo1(4).Text = "MAESTRIA" Or Combo1(4).Text = "PROFESIONAL" Then
                Text1(9).Enabled = True
                Text1(10).Enabled = True
            Else
                Text1(9) = ""
                Text1(10) = ""
                Text1(9).Enabled = False
                Text1(10).Enabled = False
            End If
    End Select
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim LenText As Long, ret As Long
    
    Select Case Index
        
        Case 0
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo1(0).hwnd, &H14C&, -1, ByVal Combo1(0).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo1(0).Text)
                     Combo1(0).ListIndex = ret
                     Combo1(0).Text = Combo1(0).List(ret)
                     Combo1(0).SelStart = LenText
                     Combo1(0).SelLength = Len(Combo1(0).Text) - LenText
                       
                  End If
            End If
            
            If Combo1(0).Text = "CASADO" Or Combo1(0).Text = "UNION LIBRE" Then
                Check5(2).Enabled = True
            Else
                Check5(2).Value = 0
                Check5(2).Enabled = False
            End If

        Case 1
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo1(1).hwnd, &H14C&, -1, ByVal Combo1(1).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo1(1).Text)
                     Combo1(1).ListIndex = ret
                     Combo1(1).Text = Combo1(1).List(ret)
                     Combo1(1).SelStart = LenText
                     Combo1(1).SelLength = Len(Combo1(1).Text) - LenText
                       
                  End If
            End If
            
        Case 2
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo1(2).hwnd, &H14C&, -1, ByVal Combo1(2).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo1(2).Text)
                     Combo1(2).ListIndex = ret
                     Combo1(2).Text = Combo1(2).List(ret)
                     Combo1(2).SelStart = LenText
                     Combo1(2).SelLength = Len(Combo1(2).Text) - LenText
                       
                  End If
            End If
            
        Case 3
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo1(3).hwnd, &H14C&, -1, ByVal Combo1(3).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo1(3).Text)
                     Combo1(3).ListIndex = ret
                     Combo1(3).Text = Combo1(3).List(ret)
                     Combo1(3).SelStart = LenText
                     Combo1(3).SelLength = Len(Combo1(3).Text) - LenText
                       
                  End If
            End If
            
        Case 4
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo1(4).hwnd, &H14C&, -1, ByVal Combo1(4).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo1(4).Text)
                     Combo1(4).ListIndex = ret
                     Combo1(4).Text = Combo1(4).List(ret)
                     Combo1(4).SelStart = LenText
                     Combo1(4).SelLength = Len(Combo1(4).Text) - LenText
                       
                  End If
            End If
            
            If Combo1(4).Text = "DOCTORADO" Or Combo1(4).Text = "MAESTRIA" Or Combo1(4).Text = "PROFESIONAL" Then
                Text1(9).Enabled = True
                Text1(10).Enabled = True
            Else
                Text1(9) = ""
                Text1(10) = ""
                Text1(9).Enabled = False
                Text1(10).Enabled = False
            End If
            
        Case 5
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo1(5).hwnd, &H14C&, -1, ByVal Combo1(5).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo1(5).Text)
                     Combo1(5).ListIndex = ret
                     Combo1(5).Text = Combo1(5).List(ret)
                     Combo1(5).SelStart = LenText
                     Combo1(5).SelLength = Len(Combo1(5).Text) - LenText
                       
                  End If
            End If
            
    End Select
    
End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Combo1(0).ListIndex = -1
        Case 1
            Cancel = Combo1(1).ListIndex = -1
        Case 2
            Cancel = Combo1(2).ListIndex = -1
        Case 3
            Cancel = Combo1(3).ListIndex = -1
        Case 4
            Cancel = Combo1(4).ListIndex = -1
        Case 5
            Cancel = Combo1(5).ListIndex = -1
    End Select
End Sub

Private Sub Combo2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim LenText As Long, ret As Long
    
    Select Case Index
        
        Case 0
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(0).hwnd, &H14C&, -1, ByVal Combo2(0).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(0).Text)
                     Combo2(0).ListIndex = ret
                     Combo2(0).Text = Combo2(0).List(ret)
                     Combo2(0).SelStart = LenText
                     Combo2(0).SelLength = Len(Combo2(0).Text) - LenText
                       
                  End If
            End If

        Case 1
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(1).hwnd, &H14C&, -1, ByVal Combo2(1).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(1).Text)
                     Combo2(1).ListIndex = ret
                     Combo2(1).Text = Combo2(1).List(ret)
                     Combo2(1).SelStart = LenText
                     Combo2(1).SelLength = Len(Combo2(1).Text) - LenText
                       
                  End If
            End If
            
        Case 2
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(2).hwnd, &H14C&, -1, ByVal Combo2(2).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(2).Text)
                     Combo2(2).ListIndex = ret
                     Combo2(2).Text = Combo2(2).List(ret)
                     Combo2(2).SelStart = LenText
                     Combo2(2).SelLength = Len(Combo2(2).Text) - LenText
                       
                  End If
            End If
            
        Case 3
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(3).hwnd, &H14C&, -1, ByVal Combo2(3).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(3).Text)
                     Combo2(3).ListIndex = ret
                     Combo2(3).Text = Combo2(3).List(ret)
                     Combo2(3).SelStart = LenText
                     Combo2(3).SelLength = Len(Combo2(3).Text) - LenText
                       
                  End If
            End If
            
        Case 4
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(4).hwnd, &H14C&, -1, ByVal Combo2(4).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(4).Text)
                     Combo2(4).ListIndex = ret
                     Combo2(4).Text = Combo2(4).List(ret)
                     Combo2(4).SelStart = LenText
                     Combo2(4).SelLength = Len(Combo2(4).Text) - LenText
                       
                  End If
            End If
            
        Case 5
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(5).hwnd, &H14C&, -1, ByVal Combo2(5).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(5).Text)
                     Combo2(5).ListIndex = ret
                     Combo2(5).Text = Combo2(5).List(ret)
                     Combo2(5).SelStart = LenText
                     Combo2(5).SelLength = Len(Combo2(5).Text) - LenText
                       
                  End If
            End If
            
        Case 6
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(6).hwnd, &H14C&, -1, ByVal Combo2(6).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(6).Text)
                     Combo2(6).ListIndex = ret
                     Combo2(6).Text = Combo2(6).List(ret)
                     Combo2(6).SelStart = LenText
                     Combo2(6).SelLength = Len(Combo2(6).Text) - LenText
                       
                  End If
            End If

        Case 7
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(7).hwnd, &H14C&, -1, ByVal Combo2(7).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(7).Text)
                     Combo2(7).ListIndex = ret
                     Combo2(7).Text = Combo2(7).List(ret)
                     Combo2(7).SelStart = LenText
                     Combo2(7).SelLength = Len(Combo2(7).Text) - LenText
                       
                  End If
            End If
            
        Case 8
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(8).hwnd, &H14C&, -1, ByVal Combo2(8).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(8).Text)
                     Combo2(8).ListIndex = ret
                     Combo2(8).Text = Combo2(8).List(ret)
                     Combo2(8).SelStart = LenText
                     Combo2(8).SelLength = Len(Combo2(8).Text) - LenText
                       
                  End If
            End If
            
        Case 9
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(9).hwnd, &H14C&, -1, ByVal Combo2(9).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(9).Text)
                     Combo2(9).ListIndex = ret
                     Combo2(9).Text = Combo2(9).List(ret)
                     Combo2(9).SelStart = LenText
                     Combo2(9).SelLength = Len(Combo2(9).Text) - LenText
                       
                  End If
            End If
            
        Case 10
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(10).hwnd, &H14C&, -1, ByVal Combo2(10).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(10).Text)
                     Combo2(10).ListIndex = ret
                     Combo2(10).Text = Combo2(10).List(ret)
                     Combo2(10).SelStart = LenText
                     Combo2(10).SelLength = Len(Combo2(10).Text) - LenText
                       
                  End If
            End If
            
        Case 11
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(11).hwnd, &H14C&, -1, ByVal Combo2(11).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(11).Text)
                     Combo2(11).ListIndex = ret
                     Combo2(11).Text = Combo2(11).List(ret)
                     Combo2(11).SelStart = LenText
                     Combo2(11).SelLength = Len(Combo2(11).Text) - LenText
                       
                  End If
            End If
            
        Case 12
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(12).hwnd, &H14C&, -1, ByVal Combo2(12).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(12).Text)
                     Combo2(12).ListIndex = ret
                     Combo2(12).Text = Combo2(12).List(ret)
                     Combo2(12).SelStart = LenText
                     Combo2(12).SelLength = Len(Combo2(12).Text) - LenText
                       
                  End If
            End If
            
        Case 13
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(13).hwnd, &H14C&, -1, ByVal Combo2(13).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(13).Text)
                     Combo2(13).ListIndex = ret
                     Combo2(13).Text = Combo2(13).List(ret)
                     Combo2(13).SelStart = LenText
                     Combo2(13).SelLength = Len(Combo2(13).Text) - LenText
                       
                  End If
            End If
            
        Case 14
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(14).hwnd, &H14C&, -1, ByVal Combo2(14).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(14).Text)
                     Combo2(14).ListIndex = ret
                     Combo2(14).Text = Combo2(14).List(ret)
                     Combo2(14).SelStart = LenText
                     Combo2(14).SelLength = Len(Combo2(14).Text) - LenText
                       
                  End If
            End If
            
        Case 15
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(15).hwnd, &H14C&, -1, ByVal Combo2(15).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(15).Text)
                     Combo2(15).ListIndex = ret
                     Combo2(15).Text = Combo2(15).List(ret)
                     Combo2(15).SelStart = LenText
                     Combo2(15).SelLength = Len(Combo2(15).Text) - LenText
                       
                  End If
            End If
            
        Case 16
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(16).hwnd, &H14C&, -1, ByVal Combo2(16).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(16).Text)
                     Combo2(16).ListIndex = ret
                     Combo2(16).Text = Combo2(16).List(ret)
                     Combo2(16).SelStart = LenText
                     Combo2(16).SelLength = Len(Combo2(16).Text) - LenText
                       
                  End If
            End If
            
        Case 17
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(17).hwnd, &H14C&, -1, ByVal Combo2(17).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(17).Text)
                     Combo2(17).ListIndex = ret
                     Combo2(17).Text = Combo2(17).List(ret)
                     Combo2(17).SelStart = LenText
                     Combo2(17).SelLength = Len(Combo2(17).Text) - LenText
                       
                  End If
            End If
            
        Case 18
            'Si los caracteres presionados están entre el 0 y la Z
            If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
              
            ret = SendMessage(Combo2(18).hwnd, &H14C&, -1, ByVal Combo2(18).Text)
              
                  If ret >= 0 Then
                     LenText = Len(Combo2(18).Text)
                     Combo2(18).ListIndex = ret
                     Combo2(18).Text = Combo2(18).List(ret)
                     Combo2(18).SelStart = LenText
                     Combo2(18).SelLength = Len(Combo2(18).Text) - LenText
                       
                  End If
            End If
            
    End Select
    
End Sub

Private Sub Combo2_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Combo2(0).ListIndex = -1
        Case 1
            Cancel = Combo2(1).ListIndex = -1
        Case 2
            Cancel = Combo2(2).ListIndex = -1
        Case 3
            Cancel = Combo2(3).ListIndex = -1
        Case 4
            Cancel = Combo2(4).ListIndex = -1
        Case 5
            Cancel = Combo2(5).ListIndex = -1
        Case 6
            Cancel = Combo2(6).ListIndex = -1
        Case 7
            Cancel = Combo2(7).ListIndex = -1
        Case 8
            Cancel = Combo2(8).ListIndex = -1
        Case 9
            Cancel = Combo2(9).ListIndex = -1
        Case 10
            Cancel = Combo2(10).ListIndex = -1
        Case 11
            Cancel = Combo2(11).ListIndex = -1
        Case 12
            Cancel = Combo2(12).ListIndex = -1
        Case 13
            Cancel = Combo2(13).ListIndex = -1
        Case 14
            Cancel = Combo2(14).ListIndex = -1
        Case 15
            Cancel = Combo2(15).ListIndex = -1
        Case 16
            Cancel = Combo2(16).ListIndex = -1
        Case 17
            Cancel = Combo2(17).ListIndex = -1
        Case 18
            Cancel = Combo2(18).ListIndex = -1
    End Select
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim LenText As Long, ret As Long

    'Si los caracteres presionados están entre el 0 y la Z
    If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
             
        ret = SendMessage(Combo3.hwnd, &H14C&, -1, ByVal Combo3.Text)
              
        If ret >= 0 Then
            LenText = Len(Combo3.Text)
            Combo3.ListIndex = ret
            Combo3.Text = Combo3.List(ret)
            Combo3.SelStart = LenText
            Combo3.SelLength = Len(Combo3.Text) - LenText
        End If
    
    End If

End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
    Cancel = Combo3.ListIndex = -1
End Sub

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        
        Case 0
            CrearTxt
        
        Case 1
            Limpiar
    
    End Select

End Sub

Private Sub Command2_Click()
    getSnapshot
End Sub

Private Sub Form_Load()
    Limpiar
    CargarCombos
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    Select Case Index
        Case 11
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 12
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 13
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
    
    End Select

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    Select Case Index
        
        Case 2
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 5
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 8
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 11
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 14
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 17
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 20
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 23
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 26
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If

        Case 29
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 32
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 35
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 38
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
        
        Case 41
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
            
        Case 44
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
            
        Case 47
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
            
        Case 50
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
            
        Case 53
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
            
        Case 56
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
    
    End Select

End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    Select Case Index
        Case 5
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
    End Select
End Sub

Private Sub Text9_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    Select Case Index
            
        Case 5
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
                If Not IsNumeric(Chr(KeyAscii)) Then
                    Beep
                    KeyAscii = 0
                End If
            End If
            
    End Select

End Sub
