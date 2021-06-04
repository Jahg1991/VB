VERSION 5.00
Begin VB.Form Form5 
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
      Left            =   9240
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   840
      Width           =   10335
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
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   7215
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

    On Error Resume Next
    
    RSCOMUNICADOS.AddNew
    
        RSCOMUNICADOS.Fields("FECHA") = Date + Time
        RSCOMUNICADOS.Fields("TITULO") = Text1.Text
        RSCOMUNICADOS.Fields("TEXTO") = Text3.Text
        
    RSCOMUNICADOS.Update
    
    msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
    
    Text1.Text = ""
    Text3.Text = ""
    
    RSCOMUNICADOS.Requery
    
    Text3.SetFocus

End Sub
