VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14520
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   14520
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   13680
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   3720
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10680
      Top             =   7440
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   4905
      Left            =   11160
      TabIndex        =   9
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8652
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12306896
      ForeColor       =   3490634
   End
   Begin VB.Image Image2 
      Height          =   4935
      Left            =   360
      Stretch         =   -1  'True
      Top             =   960
      Width           =   10455
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
      TabIndex        =   13
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
      ForeColor       =   &H0035434A&
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   240
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
      ForeColor       =   &H0035434A&
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   1
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
      ForeColor       =   &H0035434A&
      Height          =   495
      Index           =   6
      Left            =   9000
      TabIndex        =   4
      Top             =   240
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
      ForeColor       =   &H0035434A&
      Height          =   495
      Index           =   4
      Left            =   6840
      TabIndex        =   3
      Top             =   240
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
      ForeColor       =   &H0035434A&
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comunicados"
      ForeColor       =   &H0035434A&
      Height          =   375
      Index           =   3
      Left            =   11160
      TabIndex        =   8
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0035434A&
      Height          =   375
      Index           =   2
      Left            =   11160
      TabIndex        =   7
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0035434A&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nùmero de visitas:"
      ForeColor       =   &H0035434A&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Anterior"
      ForeColor       =   &H0035434A&
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Siguiente"
      ForeColor       =   &H0035434A&
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00495C67&
      Height          =   855
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   5880
      Width           =   10455
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataList1_DblClick()

    Unload Form2
    Form2.Show
    
    RSCOMUNICADOS.Filter = "TITULO like '" & DataList1.Text & "'"
    
    Form2.Caption = RSCOMUNICADOS.Fields("FECHA") + " - " + RSCOMUNICADOS.Fields("TITULO")
    
    Set Form2.Text1.DataSource = RSCOMUNICADOS
    Form2.Text1.DataField = "TEXTO"

End Sub

Private Sub Form_Load()
    
    'CARGAMOS EL RELOJ
    Label2(2).Caption = Date + Time
    
    'MOSTRAMOS EL CONTADOR
    Set Label2(1).DataSource = RSCONTADOR
    Label2(1).DataField = "CUENTA"
    
    'MOSTRAMOS COMUNICADOS
    Set DataList1.RowSource = RSCOMUNICADOS
    DataList1.ListField = "TITULO"
    
    Set Label1(0).DataSource = RSSLASH
    Label1(0).DataField = "DESCRIPCION"
    
    Set Label1(5).DataSource = RSSLASH
    Label1(5).DataField = "IMAGEN"
    
    'cargamos slash
    Image2.Picture = LoadPicture(Label1(5).Caption)
    
End Sub

Private Sub Label1_Click(Index As Integer)
    
    Select Case Index
    
        Case 1
            
            Unload Form2
            Form2.Show
                    
            Form2.Caption = "Objetivo"
                    
            Set Form2.Text1.DataSource = RSOBJETIVO
            Form2.Text1.DataField = "TEXTO"
    
        Case 2
        
            Unload Form2
            Form2.Show
        
            Form2.Caption = "Misiòn"
            
            Set Form2.Text1.DataSource = RSMISION
            Form2.Text1.DataField = "TEXTO"
        
        Case 3
            
            Unload Form2
            Form2.Show
            
            Form2.Caption = "Visiòn"
            
            Set Form2.Text1.DataSource = RSVISION
            Form2.Text1.DataField = "TEXTO"
        
        Case 4
            
            Unload Form2
            Form2.Show
        
            Form2.Caption = "Reseña"
            
            Set Form2.Text1.DataSource = RSRESENA
            Form2.Text1.DataField = "TEXTO"
        
        Case 6
            
            Form3.Caption = "Directorio"
            
            Set Form3.Label2(0).DataSource = RSAGENDA
            Form3.Label2(0).DataField = "AREA"
            
            Set Form3.Label2(1).DataSource = RSAGENDA
            Form3.Label2(1).DataField = "PUESTO"
            
            Set Form3.Label2(2).DataSource = RSAGENDA
            Form3.Label2(2).DataField = "NOMBRE"
            
            Set Form3.Label2(3).DataSource = RSAGENDA
            Form3.Label2(3).DataField = "CORREO_INTERNO"
            
            Set Form3.Label2(4).DataSource = RSAGENDA
            Form3.Label2(4).DataField = "CORREO_EXTERNO"
            
            Set Form3.Label2(5).DataSource = RSAGENDA
            Form3.Label2(5).DataField = "CELULAR"
            
            Set Form3.Label2(6).DataSource = RSAGENDA
            Form3.Label2(6).DataField = "EXTENSION"
            
            Set Form3.Label2(7).DataSource = RSAGENDA
            Form3.Label2(7).DataField = "IMAGEN"
            
            Set Form3.DataGrid1.DataSource = RSAGENDA2
            
            Form3.DataGrid1.Columns(0).Width = 4700
            
            Form3.Image1.Picture = LoadPicture(Form3.Label2(7).Caption)
            
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
            
            Form3.Show
            
    End Select
        
End Sub

Private Sub Label3_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
        
    Else
            
        If RSSLASH.BOF = True Then
                
            RSSLASH.MoveLast
                
            Image2.Picture = LoadPicture(Label1(5).Caption)
            
        Else
        
            If RSSLASH.BOF = False Then
                
                RSSLASH.MovePrevious
                
                Image2.Picture = LoadPicture(Label1(5).Caption)
                
                If Label1(5).Caption = "" Then
                
                    RSSLASH.MoveLast
                
                    Image2.Picture = LoadPicture(Label1(5).Caption)
                    
                End If
                
            End If
            
        End If
        
    End If

End Sub

Private Sub Label4_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
        
    Else
            
        If RSSLASH.EOF = True Then
                    
            RSSLASH.MoveFirst
                    
            Image2.Picture = LoadPicture(Label1(5).Caption)
                
        Else
            
            If RSSLASH.EOF = False Then
                    
                RSSLASH.MoveNext
                    
                Image2.Picture = LoadPicture(Label1(5).Caption)
                
                If Label1(5).Caption = "" Then
                
                    RSSLASH.MoveFirst
                
                    Image2.Picture = LoadPicture(Label1(5).Caption)
                    
                End If
                
            End If
                
        End If
        
    End If
    
End Sub

Private Sub Timer1_Timer()

    Label2(2).Caption = Date + Time
    
End Sub

Private Sub Timer2_Timer()
    
    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
        
    Else
            
        If RSSLASH.EOF = True Then
                    
            RSSLASH.MoveFirst
                    
            Image2.Picture = LoadPicture(Label1(5).Caption)
                
        Else
                    
            If RSSLASH.EOF = False Then
                    
                RSSLASH.MoveNext
                    
                Image2.Picture = LoadPicture(Label1(5).Caption)
                
                If Label1(5).Caption = "" Then
                
                    RSSLASH.MoveFirst
                
                    Image2.Picture = LoadPicture(Label1(5).Caption)
                    
                End If
                
            End If
                
        End If
        
    End If
    
End Sub

Private Sub Timer3_Timer()

    RSCOMUNICADOS.Requery
    
    RSSLASH.Requery
    
    'MOSTRAMOS COMUNICADOS
    Set DataList1.RowSource = RSCOMUNICADOS
    DataList1.ListField = "TITULO"

End Sub
