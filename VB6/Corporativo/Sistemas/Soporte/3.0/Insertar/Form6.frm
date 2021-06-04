VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   5565
      Left            =   240
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   10335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   5000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   240
      MaxLength       =   5000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Siguiente"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Anterior"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Form5.Show
    
End Sub

Private Sub Command2_Click()

    If RSCOMUNICADOS.EOF = True And RSCOMUNICADOS.BOF = True Then
    
    Else

        Text2.Text = Date + Time
    
        RSCOMUNICADOS.Update
        
        msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
        
    End If

End Sub

Private Sub Label1_Click()

If RSCOMUNICADOS.EOF = True And RSCOMUNICADOS.BOF = True Then
            
        Else
                
            If RSCOMUNICADOS.EOF = True Then
                    
                RSCOMUNICADOS.MoveFirst
                
            Else
            
                If RSSLASH.EOF = False Then
                    
                    RSCOMUNICADOS.MoveNext
                    
                    If Text1.Text = "" Then
                    
                        RSCOMUNICADOS.MoveFirst
                        
                    End If
                    
                End If
                
            End If
            
        End If

    
End Sub

Private Sub Label2_Click()

    If RSCOMUNICADOS.EOF = True And RSCOMUNICADOS.BOF = True Then
            
        Else
                
            If RSCOMUNICADOS.BOF = True Then
                    
                RSCOMUNICADOS.MoveLast
                
            Else
            
                If RSSLASH.BOF = False Then
                    
                    RSCOMUNICADOS.MovePrevious
                    
                    If Text1.Text = "" Then
                    
                        RSCOMUNICADOS.MoveLast
                        
                    End If
                    
                End If
                
            End If
            
        End If
    
    
End Sub
