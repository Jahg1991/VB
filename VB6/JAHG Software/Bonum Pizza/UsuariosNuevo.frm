VERSION 5.00
Begin VB.Form frmUsuariosNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Nuevo Usuario"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6915
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   5
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   4
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   3
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   2
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4215
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   6375
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
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
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contraseña"
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
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo"
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
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Articulos"
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
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ventas"
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
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Compras"
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
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inventario"
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
            Index           =   6
            Left            =   0
            TabIndex        =   11
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Corte de Caja"
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
            Index           =   7
            Left            =   0
            TabIndex        =   10
            Top             =   3600
            Width           =   1695
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
Attribute VB_Name = "frmUsuariosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To 1
        Text1(i).BackColor = &HC0C0FF
    Next i
    
    For i = 0 To 5
        With Combo1(i)
            .BackColor = &HC0C0FF
            .AddItem "Si"
            .AddItem "No"
            .Text = "No"
        End With
    Next i
    
End Sub

Private Sub Combo1_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            With Combo1(0)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 1
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 2
            With Combo1(2)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 3
            With Combo1(3)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 4
            With Combo1(4)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 5
            With Combo1(5)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
    
    End Select

End Sub

Private Sub Text1_Change(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            With Text1(0)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 1
            With Text1(1)
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

    Dim Rs As New ADODB.Recordset
    Dim In1 As Integer
    Dim i As Integer
    
    If Text1(0) <> "" And Text1(1) <> "" Then
        
        With Cn
            .CursorLocation = adUseClient
            .Open StConnection
        End With
        
        With Rs
            If .State = 1 Then .Close
                .Open "Select count(*) as existe from Usuarios where nombre like '" & Text1(0) & "';", Cn, adOpenStatic, adLockOptimistic
                .Requery
            In1 = .Fields(0).Value
            .Close
        End With
                
        If In1 = 0 Then
                    
            With Rs
                If .State = 1 Then .Close
                .Open "Select * from Usuarios;", Cn, adOpenStatic, adLockOptimistic
                .Requery
                .AddNew
                    .Fields(1).Value = Text1(0)
                    .Fields(2).Value = Text1(1)
                    .Fields(3).Value = Combo1(0)
                    .Fields(4).Value = Combo1(1)
                    .Fields(5).Value = Combo1(2)
                    .Fields(6).Value = Combo1(3)
                    .Fields(7).Value = Combo1(4)
                    .Fields(8).Value = Combo1(5)
                .Update
                .Requery
                .Close
            End With
                
            Unload frmUsuariosNuevo
            Set frmUsuariosNuevo = Nothing
            frmUsuariosNuevo.Show
                        
        Else
            MsgBox "El usuario ya existe", vbCritical, "Error"
            Text1(0).SetFocus
        End If
                
        If Cn.State = 1 Then Cn.Close
    Else
        MsgBox "Llenar todos los campos", vbCritical, "Error"
        Text1(0).SetFocus
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
