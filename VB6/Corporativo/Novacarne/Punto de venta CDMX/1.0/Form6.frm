VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13575
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5430
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   13335
         End
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    List1.Clear
        RsSaldosV.Requery
        RsSaldosV.Filter = ""
        Do Until RsSaldosV.EOF
            'JAHG Longitud Cadena
            Dim micadena As String
            Dim micadena1 As String
            Dim micadena2 As String
            Dim micadena3 As String
            Dim Saldo As Double
            Dim caracteres As Integer
            Dim caracteres1 As Integer
            micadena = Mid(RsSaldosV!Nombre, 1, 102)
            micadena1 = RsSaldosV!v1
            If IsNull(RsSaldosV!p1) = True Then
                micadena2 = "0"
            Else
                micadena2 = RsSaldosV!p1
            End If
            Saldo = micadena1 - micadena2
            micadena3 = Saldo
            caracteres = 102 - Len(micadena)
            caracteres1 = 20 - Len(micadena3)
            For i = 1 To caracteres
                micadena = micadena & " "
            Next i
            For i = 1 To caracteres1
                micadena3 = " " & micadena3
            Next i
            List1.AddItem micadena & " $" & micadena3
            RsSaldosV.MoveNext
        Loop
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Form1.Enabled = True
    
End Sub

