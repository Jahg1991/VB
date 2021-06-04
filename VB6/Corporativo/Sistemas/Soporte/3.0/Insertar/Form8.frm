VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slash"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13065
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
   ScaleHeight     =   7230
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   405
      Left            =   11760
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   6000
      Width           =   11295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   8
      Left            =   720
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   405
      Left            =   11760
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   7
      Left            =   1920
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   7320
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
      TabIndex        =   5
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
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7215
      Left            =   0
      Picture         =   "Form8.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If RSSLASH.EOF = True And RSSLASH.BOF = True Then
    
        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Imagenes|*.jpg"
        
        'Copiamos la foto
        If CommonDialog1.FileName = "" Then
        
        Else
            Text2(1).Text = 1
            FileCopy CommonDialog1.FileName, "C:\JAHG Software\JAHG Soporte\Slash\" + Text2(1).Text + ".jpg"
            Text2(7).Text = "C:\JAHG Software\JAHG Soporte\Slash\" + Text2(1).Text + ".jpg"
            
            Image1.Picture = LoadPicture(Text2(7).Text)
    
    Else

        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Imagenes|*.jpg"
        
        'Copiamos la foto
        If CommonDialog1.FileName = "" Then
        
        Else
            Text2(1).Text = Text2(8).Text + 1
            FileCopy CommonDialog1.FileName, "C:\JAHG Software\JAHG Soporte\Slash\" + Text2(1).Text + ".jpg"
            Text2(7).Text = "C:\JAHG Software\JAHG Soporte\Slash\" + Text2(1).Text + ".jpg"
            
            Image1.Picture = LoadPicture(Text2(7).Text)
            
        End If
        
    End If
    
End Sub

Private Sub Command2_Click()

    On Error Resume Next

    RSSLASH.AddNew
    
        RSSLASH.Fields("IMAGEN") = Text2(7).Text
        RSSLASH.Fields("DESCRIPCION") = Text2(0).Text
        
    RSSLASH.Update
    
    msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
    
    RSSLASH.Requery
    
    RSSLASHID.Requery
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(7).Text = ""
    
    Image1.Picture = LoadPicture(Text2(7).Text)

End Sub
    
    
