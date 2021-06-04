VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soporte"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   5880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Slash"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comunicados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Objetivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Misiòn"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reseña"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visiòn"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   5760
      Left            =   0
      Picture         =   "Form1.frx":1EDAA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    'CARGAMOS EL RELOJ
    Label2(2).Caption = Date + Time
    
End Sub

Private Sub Label1_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
        
            Unload Form6
            Form6.Show
                    
            Form6.Caption = "Comunicados"
                    
            Set Form6.Text1.DataSource = RSCOMUNICADOS
            Form6.Text1.DataField = "TITULO"
            
            Set Form6.Text2.DataSource = RSCOMUNICADOS
            Form6.Text2.DataField = "FECHA"
            
            Set Form6.Text3.DataSource = RSCOMUNICADOS
            Form6.Text3.DataField = "TEXTO"
    
        Case 1
            
            Unload Form2
            Form2.Show
                    
            Form2.Caption = "Objetivo"
                    
            Set Form2.Text1.DataSource = RSOBJETIVO
            Form2.Text1.DataField = "TEXTO"
            
            Form2.Command1(0).Visible = False
            Form2.Command1(1).Visible = True
            Form2.Command1(2).Visible = False
            Form2.Command1(3).Visible = False
    
        Case 2
        
            Unload Form2
            Form2.Show
        
            Form2.Caption = "Misiòn"
            
            Set Form2.Text1.DataSource = RSMISION
            Form2.Text1.DataField = "TEXTO"
            
            Form2.Command1(0).Visible = True
            Form2.Command1(1).Visible = False
            Form2.Command1(2).Visible = False
            Form2.Command1(3).Visible = False
        
        Case 3
            
            Unload Form2
            Form2.Show
            
            Form2.Caption = "Visiòn"
            
            Set Form2.Text1.DataSource = RSVISION
            Form2.Text1.DataField = "TEXTO"
            
            Form2.Command1(0).Visible = False
            Form2.Command1(1).Visible = False
            Form2.Command1(2).Visible = True
            Form2.Command1(3).Visible = False
        
        Case 4
            
            Unload Form2
            Form2.Show
        
            Form2.Caption = "Reseña"
            
            Set Form2.Text1.DataSource = RSRESENA
            Form2.Text1.DataField = "TEXTO"
            
            Form2.Command1(0).Visible = False
            Form2.Command1(1).Visible = False
            Form2.Command1(2).Visible = False
            Form2.Command1(3).Visible = True
        
        Case 6
            
            Form3.Caption = "Directorio"
            
            RSAGENDA.Requery
            
            RSAGENDA.Requery
            
            Set Form3.Text2(0).DataSource = RSAGENDA
            Form3.Text2(0).DataField = "AREA"
            
            Set Form3.Text2(1).DataSource = RSAGENDA
            Form3.Text2(1).DataField = "PUESTO"
            
            Set Form3.Text2(2).DataSource = RSAGENDA
            Form3.Text2(2).DataField = "NOMBRE"
            
            Set Form3.Text2(3).DataSource = RSAGENDA
            Form3.Text2(3).DataField = "CORREO_INTERNO"
            
            Set Form3.Text2(4).DataSource = RSAGENDA
            Form3.Text2(4).DataField = "CORREO_EXTERNO"
            
            Set Form3.Text2(5).DataSource = RSAGENDA
            Form3.Text2(5).DataField = "CELULAR"
            
            Set Form3.Text2(6).DataSource = RSAGENDA
            Form3.Text2(6).DataField = "EXTENSION"
            
            Set Form3.Text2(7).DataSource = RSAGENDA
            Form3.Text2(7).DataField = "IMAGEN"
            
            Set Form3.Text2(8).DataSource = RSAGENDA
            Form3.Text2(8).DataField = "ID"
            
            Set Form3.DataGrid1.DataSource = RSAGENDA2
            
            Form3.DataGrid1.Columns(0).Width = 4700
            
            Form3.Image1.Picture = LoadPicture(Form3.Text2(7).Text)
            
            Form3.Show
            
        Case 7
        
            Set Form7.Text2(8).DataSource = RSSLASH
            Form7.Text2(8).DataField = "ID"
            
            Set Form7.Text2(7).DataSource = RSSLASH
            Form7.Text2(7).DataField = "IMAGEN"
            
            Set Form7.Text2(0).DataSource = RSSLASH
            Form7.Text2(0).DataField = "DESCRIPCION"
            
            Form7.Image1.Picture = LoadPicture(Form7.Text2(7).Text)
            
            Form7.Show
            
    End Select
        
End Sub

Private Sub Timer1_Timer()

    Label2(2).Caption = Date + Time
    
End Sub

