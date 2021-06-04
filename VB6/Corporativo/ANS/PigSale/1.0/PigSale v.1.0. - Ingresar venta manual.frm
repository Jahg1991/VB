VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PigSale v.1.0 - Venta manual"
   ClientHeight    =   9510
   ClientLeft      =   7650
   ClientTop       =   390
   ClientWidth     =   5190
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   5190
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   510
      Index           =   8
      Left            =   1320
      TabIndex        =   17
      Top             =   9720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   510
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   9720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   5
         Left            =   4320
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   4
         Left            =   1680
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":0681
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   8640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   3
         Left            =   1680
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   0
         Left            =   1680
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":16C2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7080
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   2
         Left            =   1680
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":1F89
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   6
         Left            =   1800
         TabIndex        =   10
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   5
         Left            =   1800
         TabIndex        =   9
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   4
         Left            =   1800
         TabIndex        =   8
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   3
         Left            =   1800
         TabIndex        =   7
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   510
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   3
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Format          =   97452033
         CurrentDate     =   42394
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   8
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":27BD
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   7
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":3135
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   6
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":3AC4
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   5
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":42BF
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   4
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":4950
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   3
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":5212
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   2
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":5C57
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":67E3
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "PigSale v.1.0. - Ingresar venta manual.frx":706D
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   9720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
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
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            Form4.Show
            Form4.Label1.Caption = Label1.Caption
            Unload Me
        
        Case 2
            With RsSales
                .Requery
                .AddNew
                .Fields("FECHA") = DTPicker1.Value
                .Fields("GRANJA") = Combo1.Text
                .Fields("NUMERO") = Text1(0).Text
                .Fields("KILOS") = Text1(1).Text
                .Fields("PROMEDIO") = Text1(2).Text
                .Fields("CLIENTE") = Text1(3).Text
                .Fields("TEJABAN") = Text1(4).Text
                .Fields("MORTANDAD") = Text1(5).Text
                .Fields("OBSERVACIONES") = Text1(6).Text
                .Fields("ANO") = Text1(7).Text
                .Fields("SEMANA") = Text1(8).Text
                .Update
                .Requery
            End With
            
            DTPicker1.Value = Date
            Combo1.Text = ""
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
            Text1(5).Text = ""
            Text1(6).Text = ""
                
            DTPicker1.SetFocus
            
        Case 3
            Form2.Show
            Form2.Label1.Caption = Label1.Caption
            Unload Me
            
        Case 4
            Unload Me
        
        Case 5
            Form1.Show
            Form1.Text1(0).Text = ""
            Form1.Text1(1).Text = ""
            Unload Me
            
    End Select
    
End Sub

Private Sub Form_Load()

    On Error Resume Next

    With RsSales
        If .State = 1 Then .Close
           .Open "Select * from VENTAS", CnDb, adOpenStatic, adLockOptimistic
           .Requery
    End With

    DTPicker1.Value = Date
    
    With Combo1
        .AddItem ("El Moro")
        .AddItem ("El Terrero")
        .AddItem ("Isabel")
        .AddItem ("La Cuna")
        .AddItem ("La Laja")
        .AddItem ("La Loma")
        .AddItem ("San Aparicio")
    End With
    
    With RSYearWeek
        If .State = 1 Then .Close
           .Open "SELECT DATEPART( yyyy, GETDATE() ) Ano,DATEPART( wk, GETDATE() ) Semana", CnDb, adOpenStatic, adLockOptimistic
           .Requery
    End With
    
   Set DataGrid1.DataSource = RSYearWeek
   Text1(7) = DataGrid1.Columns(0).Value
   Text1(8) = DataGrid1.Columns(1).Value
    
End Sub

Private Sub Text1_Change(Index As Integer)

    On Error Resume Next
    
    Select Case Index
        
        Case 0
            Text1(2) = Round(Text1(1) / Text1(0), 2)
        
        Case 1
            Text1(2) = Round(Text1(1) / Text1(0), 2)
        
    End Select

End Sub
