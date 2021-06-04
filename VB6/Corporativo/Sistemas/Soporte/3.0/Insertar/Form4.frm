VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7320
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
   ScaleHeight     =   7080
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   0
      Left            =   2280
      TabIndex        =   18
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   1
      Left            =   2280
      TabIndex        =   17
      Top             =   2880
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   2
      Left            =   2280
      TabIndex        =   16
      Top             =   3480
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   3
      Left            =   2280
      TabIndex        =   15
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   4
      Left            =   2280
      TabIndex        =   14
      Top             =   4680
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   5280
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   6
      Left            =   2280
      TabIndex        =   12
      Top             =   5880
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   405
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   9
      Left            =   1800
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   8
      Left            =   1320
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   405
      Left            =   6000
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
      Left            =   720
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Àrea"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Puesto"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correo interno"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correo externo"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nùmero de celular"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Extensiòn"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Left            =   0
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If RSAGENDA.BOF = True And RSAGENDA.EOF And RSAGENDA2.BOF = True And RSAGENDA2.EOF = True Then
    
        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Imagenes|*.jpg"
        
        'Copiamos la foto
        If CommonDialog1.FileName = "" Then
        
        Else
            Text2(9).Text = 1
            FileCopy CommonDialog1.FileName, "C:\JAHG Software\JAHG Soporte\Empleados\Imagenes\" + Text2(9).Text + ".jpg"
            Text2(7).Text = "C:\JAHG Software\JAHG Soporte\Empleados\Imagenes\" + Text2(9).Text + ".jpg"
            
            Image1.Picture = LoadPicture(Text2(7).Text)
            
            Text2(0).SetFocus
            
        End If
    
    Else

        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Imagenes|*.jpg"
        
        'Copiamos la foto
        If CommonDialog1.FileName = "" Then
        
        Else
            Text2(9).Text = Text2(8).Text + 1
            FileCopy CommonDialog1.FileName, "C:\JAHG Software\JAHG Soporte\Empleados\Imagenes\" + Text2(9).Text + ".jpg"
            Text2(7).Text = "C:\JAHG Software\JAHG Soporte\Empleados\Imagenes\" + Text2(9).Text + ".jpg"
            
            Image1.Picture = LoadPicture(Text2(7).Text)
            
            Text2(0).SetFocus
            
        End If
        
    End If
    
End Sub

Private Sub Command2_Click()

    On Error Resume Next
    
    RSAGENDA.AddNew
    
        RSAGENDA.Fields("AREA") = Text2(0).Text
        RSAGENDA.Fields("PUESTO") = Text2(1).Text
        RSAGENDA.Fields("NOMBRE") = Text2(2).Text
        RSAGENDA.Fields("CORREO_INTERNO") = Text2(3).Text
        RSAGENDA.Fields("CORREO_EXTERNO") = Text2(4).Text
        RSAGENDA.Fields("CELULAR") = Text2(5).Text
        RSAGENDA.Fields("EXTENSION") = Text2(6).Text
        RSAGENDA.Fields("IMAGEN") = Text2(7).Text
        
    RSAGENDA.Update
    
    msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Text2(3).Text = ""
    Text2(4).Text = ""
    Text2(5).Text = ""
    Text2(6).Text = ""
    Text2(7).Text = ""
    Text2(9).Text = ""
    
    Image1.Picture = LoadPicture(Text2(7).Text)
    
    RSID.Requery
    
    RSAGENDA.Requery
    
    RSAGENDA2.Requery
    
    Text2(0).SetFocus
    
End Sub
