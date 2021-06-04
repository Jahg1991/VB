VERSION 5.00
Begin VB.Form frmEntradaSalidaDinero 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Movimientos de efectivo"
   ClientHeight    =   1785
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   7455
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
   ScaleHeight     =   1785
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1560
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1335
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   6975
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   1560
            TabIndex        =   0
            Top             =   120
            Width           =   5175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   1
            Left            =   1560
            MaxLength       =   250
            TabIndex        =   1
            Top             =   720
            Width           =   5175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   -600
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comentario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   -600
            TabIndex        =   4
            Top             =   720
            Width           =   2055
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmEntradaSalidaDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset

Private Sub Form_Load()

    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs
                
        If .State = 1 Then .Close
        .Open "Select * from MovimientosCaja", Cn, adOpenStatic, adLockOptimistic
        .Requery
    
    End With
    
    Text2(0).BackColor = &HC0C0FF
    Text2(1).BackColor = &HC0C0FF

End Sub

Private Sub Text2_Change(Index As Integer)

    On Error Resume Next
    
    Select Case Index
        
        Case 0
            With Text2(0)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 1
            With Text2(1)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
            
    End Select
    
End Sub

Private Sub Guardar_Click()

    On Error Resume Next
    
    If Val(Text2(0)) <> 0 Then
            
        With Rs
                
            .AddNew
            .Fields(1) = Date
                    
            If StTipoEntradaSalida = "Entrada" Then
                .Fields(2) = "Entrada Manual de Efectivo"
                .Fields(3) = Replace(Format(Val(Text2(0)), "0.00"), ",", ".")
            End If
                    
            If StTipoEntradaSalida = "Salida" Then
                .Fields(2) = "Retiro Manual de Efectivo"
                .Fields(3) = Replace(Format(Val(Text2(0) * -1), "0.00"), ",", ".")
            End If
                    
            .Fields(4) = Text2(1)
            .Fields(5) = "No"
                    
            .Update
            .Requery
                
        End With
        
        frmMenuInicial.Enabled = True
        Unload Me
    
    Else
    
        MsgBox "Monto no válido", vbCritical, "Advertencia"
    
    End If

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next

    frmMenuInicial.Enabled = True
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Sub

