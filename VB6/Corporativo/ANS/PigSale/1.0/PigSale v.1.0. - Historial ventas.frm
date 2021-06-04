VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PigSale v.1.0 - Historial de ventas"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13485
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
   ScaleHeight     =   7230
   ScaleWidth      =   13485
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Height          =   615
      Index           =   6
      Left            =   4440
      Picture         =   "PigSale v.1.0. - Historial ventas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   9
         Left            =   12600
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":08C7
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   8
         Left            =   7440
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":0F48
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   7
         Left            =   5880
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":15D3
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   5
         Left            =   11640
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":1F89
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   4
         Left            =   10080
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":285A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   2
         Left            =   4920
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":308E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   1
         Left            =   3360
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":3856
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   0
         Left            =   1680
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":40F2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   3
         Left            =   120
         Picture         =   "PigSale v.1.0. - Historial ventas.frx":4920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
         Caption         =   "Historial de ventas"
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
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            With RsSales
                .MovePrevious
            End With
        
        Case 1
            With RsSales
                .MoveNext
            End With
        
        Case 2
            With RsSales
                .MoveLast
            End With
        
        Case 3
            With RsSales
                .MoveFirst
            End With
        
        Case 4
            With RsSales
                .Requery
                .Update
                .Requery
            End With
        
        Case 5
            With RsSales
                .Delete
            End With
        
        Case 6
            Form3.Show
            Form3.Label1.Caption = Label1.Caption
            Unload Me
        
        Case 7
            Form2.Show
            Form2.Label1.Caption = Label1.Caption
            Unload Me
        
        Case 8
            Unload Me
        
        Case 9
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
    
    Set DataGrid1.DataSource = RsSales
    
End Sub
