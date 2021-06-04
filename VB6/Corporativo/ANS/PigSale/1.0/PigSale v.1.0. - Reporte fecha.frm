VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PigSale v.1.0 - Reporte Fecha"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3885
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
   ScaleHeight     =   9360
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2143
      _Version        =   393216
      Enabled         =   -1  'True
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   5
         Left            =   3000
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   4
         Left            =   1080
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":0681
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   3
         Left            =   1080
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   2
         Left            =   1080
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":16C2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   1
         Left            =   1920
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":1F89
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   825
         Width           =   855
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         Max             =   52
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   0
         Left            =   240
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":29C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   1455
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2013
         Max             =   2050
         Min             =   2013
         Enabled         =   -1  'True
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   120
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":3460
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "PigSale v.1.0. - Reporte fecha.frx":3C91
         Stretch         =   -1  'True
         Top             =   720
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
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2143
      _Version        =   393216
      Enabled         =   -1  'True
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2143
      _Version        =   393216
      Enabled         =   -1  'True
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
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0

            
        Case 1
            
    
        Case 2

            
        Case 3

            
        Case 4

        
        Case 5


    End Select
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
            
    With RSReports
        If .State = 1 Then .Close
            .Open "Select FECHA, GRANJA, NUMERO, KILOS, PROMEDIO, CLIENTE, TEJABAN, MORTANDAD, OBSERVACIONES from VENTAS", CnDb, adOpenStatic, adLockOptimistic
            .Requery
    End With
            
    With RSTotals
            If .State = 1 Then .Close
            .Open "Select Sum(NUMERO) TOTNUM, Sum(KILOS) TOTKIL, (TOTKIL/TOTNUM) PROMEDIO", CnDb, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSYearWeek
            If .State = 1 Then .Close
            .Open "Select DATEPART( yyyy, GETDATE() ) ANO, DATEPART( wk, GETDATE() ) SEMANA", CnDb, adOpenStatic, adLockOptimistic
            .Requery
    End With
        
    Set DataGrid1(0).DataSource = RSReports
    Set DataGrid1(1).DataSource = RSTotals
    Set DataGrid1(2).DataSource = RSYearWeek
    
    UpDown1(0).Value = DataGrid1(2).Columns(1).Value
    UpDown1(1).Value = DataGrid1(2).Columns(0).Value
    
    Text1(0).Text = UpDown1(0).Value
    Text1(1).Text = UpDown1(1).Value
    
End Sub

Private Sub Text1_Change(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            With RSReports
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "SEMANA LIKE '*" & Text1(0) & "*'"
                Else
                    .Filter = ""
                    Set DataGrid1(0).DataSource = RSReports
                    .MoveFirst
                End If
            End With
            With RSTotals
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "SEMANA LIKE '*" & Text1(0) & "*'"
                Else
                    .Filter = ""
                    Set DataGrid1(1).DataSource = RSTotals
                    .MoveFirst
                End If
            End With
            
        Case 1
            With RSReports
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "SEMANA LIKE '*" & Text1(0) & "*'"
                Else
                    .Filter = ""
                    Set DataGrid1(0).DataSource = RSReports
                    .MoveFirst
                End If
            End With
            With RSTotals
                .Requery
                If OPTION1.Value = True Then
                    .Filter = "SEMANA LIKE '*" & Text1(0) & "*'"
                Else
                    .Filter = ""
                    Set DataGrid1(1).DataSource = RSTotals
                    .MoveFirst
                End If
            End With
            
    End Select
    
End Sub

Private Sub UpDown1_DownClick(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            Text1(0).Text = UpDown1(0).Value
            
        Case 1
            Text1(1).Text = UpDown1(1).Value
            
    End Select

End Sub

Private Sub UpDown1_UpClick(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            Text1(0).Text = UpDown1(0).Value
            
        Case 1
            Text1(1).Text = UpDown1(1).Value
            
    End Select
    
End Sub
