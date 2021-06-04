VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Empresa"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   7305
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
   Begin VB.Menu Abrir 
      Caption         =   "Abrir"
   End
   Begin VB.Menu Nuevo 
      Caption         =   "Nuevo"
   End
   Begin VB.Menu Editar 
      Caption         =   "Editar"
   End
   Begin VB.Menu Eliminar 
      Caption         =   "Eliminar"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abrir_Click()

    With RsEmpresaAbierta
    
        .Filter = "Nombre = '" & List1.Text & "'"
        
        EmpAbierta = RsEmpresaAbierta.Fields("Id")
    
    End With

End Sub

Private Sub Editar_Click()

    With RsEmpresaAbierta
    
        .Filter = "Nombre = '" & List1.Text & "'"
        
        EmpAbierta = RsEmpresaAbierta.Fields("Id")
    
    End With

    CTipo = 3
    
    GTipo = 2
    
    InitForm
    
    Unload Me

End Sub

Private Sub Eliminar_Click()

    With RsEmpresaAbierta
    
        .Filter = "Nombre = '" & List1.Text & "'"
        
        .Delete
        
        .Update
    
    End With
    
    CTipo = 1
    
    InitForm

End Sub

Private Sub Nuevo_Click()

    CTipo = 2
    
    GTipo = 1
    
    InitForm
    
    Unload Me

End Sub
