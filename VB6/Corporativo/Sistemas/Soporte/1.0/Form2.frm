VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imàgen adjunta"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14490
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   9375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Image1.Picture = LoadPicture(Form1.RutaOrigen)
    
End Sub
