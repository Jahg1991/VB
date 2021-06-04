VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ClipControls    =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1440
      Top             =   2400
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   720
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
With Form1
    .Height = Screen.Height - 600
    .Width = Screen.Width
End With
With Image1
    On Error GoTo err
    .Picture = LoadPicture(App.Path & "\Fondo.jpg")
    .Top = 0
    .Left = 0
End With
Exit Sub
err:
    Unload Me
End Sub

Private Sub Form_Resize()
With Image1
    .Width = Form1.Width
    .Height = Form1.Height
End With
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
With Form1
    .Height = Screen.Height - 600
    .Width = Screen.Width
End With
With Image1
    .Width = Form1.Width
    .Height = Form1.Height
End With
With Image1
    .Picture = LoadPicture(App.Path & "\Fondo.jpg")
End With
End Sub
