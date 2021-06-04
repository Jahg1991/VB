VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slash"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
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
   ScaleHeight     =   7065
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   5880
      Width           =   11295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   405
      Left            =   11760
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   8
      Left            =   0
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   405
      Left            =   9120
      TabIndex        =   3
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nuevo"
      Height          =   405
      Left            =   11760
      TabIndex        =   2
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   405
      Left            =   10440
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   7
      Left            =   600
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   12615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Anterior"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Siguiente"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Left            =   0
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
    
    Else

        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Imagenes|*.jpg"
        
        'Copiamos la foto
        If CommonDialog1.FileName = "" Then
        
        Else
            FileCopy CommonDialog1.FileName, "C:\JAHG Software\JAHG Soporte\Slash\" + Text2(8).Text + ".jpg"
            Text2(7).Text = "C:\JAHG Software\JAHG Soporte\Slash\" + Text2(8).Text + ".jpg"
            
            Image1.Picture = LoadPicture(Text2(7).Text)
            
        End If
        
    End If
    
End Sub

Private Sub Command2_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
    
    Else

        RSSLASH.Update
        
        msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
        
        RSSLASH.Update
        
        Image1.Picture = LoadPicture(Text2(7).Text)
        
    End If

End Sub

Private Sub Command3_Click()
    
    Form8.Show
    
    Set Form8.Text2(8).DataSource = RSSLASHID
    Form8.Text2(8).DataField = "ID"
    
End Sub

Private Sub Command4_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
    
    Else

        RSSLASH.Delete
        
        msg = MsgBox("Eliminado correctamente", vbOKOnly, "Listo!")
        
        RSSLASH.Requery
        
        Image1.Picture = LoadPicture(Text2(7).Text)
        
    End If

End Sub

Private Sub Label1_Click(Index As Integer)

    Select Case Index
    
        Case 1
    
            If RSSLASH.EOF = True And RSSLASH.BOF = True Then
                
            Else
                    
                If RSSLASH.EOF = True Then
                        
                    RSSLASH.MoveLast
                    
                    Image1.Picture = LoadPicture(Text2(7).Text)
                    
                Else
                
                    If RSSLASH.EOF = False Then
                        
                        RSSLASH.MovePrevious
                        
                        Image1.Picture = LoadPicture(Text2(7).Text)
                        
                        If Text2(8).Text = "" Then
                        
                            RSSLASH.MoveLast
                            
                            Image1.Picture = LoadPicture(Text2(7).Text)
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
        
    End Select

End Sub

Private Sub Label2_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
            
    Else
                
        If RSSLASH.BOF = True Then
                    
            RSSLASH.MoveFirst
            
            Image1.Picture = LoadPicture(Text2(7).Text)
                
        Else
            
            If RSSLASH.BOF = False Then
                    
                RSSLASH.MoveNext
                
                Image1.Picture = LoadPicture(Text2(7).Text)
                    
                If Text2(8).Text = "" Then
                    
                    RSSLASH.MoveFirst
                    
                    Image1.Picture = LoadPicture(Text2(7).Text)
                        
                End If
                    
            End If
                
        End If
        
    End If
            
End Sub
    
    
