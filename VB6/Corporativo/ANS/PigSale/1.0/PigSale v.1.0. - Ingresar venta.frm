VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PigSale v.1.0 - Nueva venta"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
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
   ScaleHeight     =   4830
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Height          =   615
      Index           =   2
      Left            =   960
      Picture         =   "PigSale v.1.0. - Ingresar venta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Index           =   1
      Left            =   960
      Picture         =   "PigSale v.1.0. - Ingresar venta.frx":08C7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   5
         Left            =   2520
         Picture         =   "PigSale v.1.0. - Ingresar venta.frx":1322
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   4
         Left            =   840
         Picture         =   "PigSale v.1.0. - Ingresar venta.frx":19A3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   3
         Left            =   840
         Picture         =   "PigSale v.1.0. - Ingresar venta.frx":202E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   0
         Left            =   840
         Picture         =   "PigSale v.1.0. - Ingresar venta.frx":29E4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            Form5.Show
            Form5.Label1.Caption = Label1.Caption
            Unload Me
        
        Case 1
            Form6.Show
            Form6.Label1.Caption = Label1.Caption
            Unload Me
    
        Case 2
            Form3.Show
            Form3.Label1.Caption = Label1.Caption
            Unload Me
            
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
