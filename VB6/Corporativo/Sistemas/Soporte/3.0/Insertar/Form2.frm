VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
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
   ScaleHeight     =   7770
   ScaleWidth      =   10800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   3
      Left            =   9240
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   2
      Left            =   9240
      TabIndex        =   3
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   6525
      Left            =   240
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    Select Case Index
    
        Case 0
        
            RSMISION.Update
            
            msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
        
        Case 1
        
            RSOBJETIVO.Update
            
            msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
            
        Case 2
        
            RSVISION.Update
            
            msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
        
        Case 3
        
            RSRESENA.Update
            
            msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
        
    End Select
    
End Sub
