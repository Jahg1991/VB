VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13635
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12306896
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
      ForeColor       =   3490634
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBC9D0&
      BorderStyle     =   0  'None
      ForeColor       =   &H0035434A&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BBC9D0&
      BorderStyle     =   0  'None
      Caption         =   "Ficha"
      ForeColor       =   &H0035434A&
      Height          =   6615
      Left            =   6000
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.CommandButton Command2 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   15
         Top             =   6000
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   5
         Left            =   2280
         TabIndex        =   14
         Top             =   5400
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   13
         Top             =   4800
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   12
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Top             =   3600
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Extensiòn"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nùmero de celular"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Correo externo"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Correo interno"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puesto"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Àrea"
         ForeColor       =   &H0035434A&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.Image Image2 
      Height          =   7215
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Form4.Show
    Form4.txt_Asunto.SetFocus

End Sub

Private Sub Command2_Click()

    Form5.Show
    Form5.txt_Asunto.SetFocus

End Sub

Private Sub DataGrid1_Click()

    Text1.Text = DataGrid1.Columns(0).Text

End Sub

Private Sub Text1_Change()

    DataGrid1.Columns(0).Width = 4700
    
    RSAGENDA.Requery
    
    If Text1 = "" Then
    
        RSAGENDA.Filter = ""
        RSAGENDA2.Filter = ""
        
        DataGrid1.Columns(0).Width = 4700
        
        Image1.Picture = LoadPicture(Label2(7).Caption)
        
        If RSAGENDA.BOF = True And RSAGENDA.EOF And RSAGENDA2.BOF = True And RSAGENDA2.EOF = True Then
        
        Else
        
            RSAGENDA.MoveFirst
            RSAGENDA2.MoveFirst
            
                
            Set Label2(0).DataSource = RSAGENDA
            Label2(0).DataField = "AREA"
                    
            Set Label2(1).DataSource = RSAGENDA
            Label2(1).DataField = "PUESTO"
                    
            Set Label2(2).DataSource = RSAGENDA
            Label2(2).DataField = "NOMBRE"
                    
            Set Label2(3).DataSource = RSAGENDA
            Label2(3).DataField = "CORREO_INTERNO"
                    
            Set Label2(4).DataSource = RSAGENDA
            Label2(4).DataField = "CORREO_EXTERNO"
                    
            Set Label2(5).DataSource = RSAGENDA
            Label2(5).DataField = "CELULAR"
                    
            Set Label2(6).DataSource = RSAGENDA
            Label2(6).DataField = "EXTENSION"
                   
            Set Label2(7).DataSource = RSAGENDA
            Label2(7).DataField = "IMAGEN"
                    
            Set DataGrid1.DataSource = RSAGENDA2
            
            DataGrid1.Columns(0).Width = 4700
            
            Image1.Picture = LoadPicture(Label2(7).Caption)
            
            If Form3.Label2(3).Caption = "" Then
            
                Form3.Command1.Enabled = False
            
            Else
            
                Form3.Command1.Enabled = True
            
            End If
            
            If Form3.Label2(4).Caption = "" Then
            
                Form3.Command2.Enabled = False
            
            Else
            
                Form3.Command2.Enabled = True
            
            End If
            
        End If
    
    Else
    
        RSAGENDA.Filter = "NOMBRE like '*" & Text1 & "*'"
        RSAGENDA2.Filter = "NOMBRE like '*" & Text1 & "*'"
            
        Set Label2(0).DataSource = RSAGENDA
        Label2(0).DataField = "AREA"
                
        Set Label2(1).DataSource = RSAGENDA
        Label2(1).DataField = "PUESTO"
                
        Set Label2(2).DataSource = RSAGENDA
        Label2(2).DataField = "NOMBRE"
                
        Set Label2(3).DataSource = RSAGENDA
        Label2(3).DataField = "CORREO_INTERNO"
                
        Set Label2(4).DataSource = RSAGENDA
        Label2(4).DataField = "CORREO_EXTERNO"
                
        Set Label2(5).DataSource = RSAGENDA
        Label2(5).DataField = "CELULAR"
                
        Set Label2(6).DataSource = RSAGENDA
        Label2(6).DataField = "EXTENSION"
               
        Set Label2(7).DataSource = RSAGENDA
        Label2(7).DataField = "IMAGEN"
                
        Set DataGrid1.DataSource = RSAGENDA2
        
        DataGrid1.Columns(0).Width = 4700
        
        Image1.Picture = LoadPicture(Label2(7).Caption)
        
        If Form3.Label2(3).Caption = "" Then
            
            Form3.Command1.Enabled = False
            
        Else
            
            Form3.Command1.Enabled = True
            
        End If
            
        If Form3.Label2(4).Caption = "" Then
            
            Form3.Command2.Enabled = False
            
        Else
            
            Form3.Command2.Enabled = True
            
        End If
        
    End If
    
End Sub
