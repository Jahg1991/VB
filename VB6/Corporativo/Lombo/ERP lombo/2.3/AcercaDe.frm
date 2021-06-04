VERSION 5.00
Begin VB.Form frmAcercaDe 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Acerca de"
   ClientHeight    =   3615
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   6900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   63.765
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   121.708
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6660
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6375
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PUNTO DE VENTA   NÚMERO DE SOPORTE                   (332) 080 2351    JUAN ALFREDO HERNÁNDEZ GONZÁLEZ           2020"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   3360
            TabIndex        =   2
            Top             =   360
            Width           =   2775
         End
         Begin VB.Image Image1 
            Height          =   2500
            Left            =   360
            Picture         =   "AcercaDe.frx":0000
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2500
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAcercaDe:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
