VERSION 5.00
Begin VB.Form frmClientesProveedoresNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Añadir Clientes / Proveedores"
   ClientHeight    =   4680
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   10410
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
   ScaleHeight     =   4680
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10140
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   4215
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9855
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   7
            Left            =   2280
            TabIndex        =   17
            Top             =   3600
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   6
            Left            =   2280
            TabIndex        =   16
            Top             =   3120
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   5
            Left            =   2280
            TabIndex        =   15
            Top             =   2640
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   4
            Left            =   2280
            TabIndex        =   14
            Top             =   2160
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   3
            Left            =   2280
            TabIndex        =   13
            Top             =   1680
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   2
            Left            =   2280
            TabIndex        =   12
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   1
            Left            =   2280
            TabIndex        =   11
            Top             =   720
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   0
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Width           =   7335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencias"
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
            Left            =   360
            TabIndex        =   9
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Correo Electronico"
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
            TabIndex        =   8
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
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
            Left            =   720
            TabIndex        =   7
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Postal"
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
            Left            =   480
            TabIndex        =   6
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Colonia"
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
            Left            =   720
            TabIndex        =   5
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   720
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calle"
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
            Left            =   720
            TabIndex        =   3
            Top             =   720
            Width           =   1455
         End
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
            Left            =   720
            TabIndex        =   2
            Top             =   240
            Width           =   1455
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
Attribute VB_Name = "frmClientesProveedoresNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To 7
        Text1(i).BackColor = &HC0C0FF
    Next i

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
        
        Case 2
            With Text1(2)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 3
            With Text1(3)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 4
            With Text1(4)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 5
            With Text1(5)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 6
            With Text1(6)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 7
            With Text1(7)
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
    
    Dim In1 As Integer
    Dim i As Integer
    
    If Text1(0) <> "" Then
        
        With Cn
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            .Open StConnection
        End With
        
        With Rs
            If .State = 1 Then .Close
                .Open "Select count(*) as existe from ClientesProveedores where nombre like '" & Text1(0) & "';", Cn, adOpenStatic, adLockOptimistic
                .Requery
            In1 = .Fields(0).Value
            .Close
        End With
        
        If In1 = 0 Then
            
            With Rs
                If .State = 1 Then .Close
                .Open "Select * from ClientesProveedores;", Cn, adOpenStatic, adLockOptimistic
                .Requery
                .AddNew
                    .Fields(1).Value = Text1(0)
                    .Fields(2).Value = Text1(1)
                    .Fields(3).Value = Text1(2)
                    .Fields(4).Value = Text1(3)
                    .Fields(5).Value = Text1(4)
                    .Fields(6).Value = Text1(5)
                    .Fields(7).Value = Text1(6)
                    .Fields(8).Value = Text1(7)
                    .Fields(9).Value = StTipoClienteProveedor
                .Update
                .Requery
                .Close
            End With
            
            If InTipoAltaClienteProveedor = 1 Then
                
                frmVentasCompras.Enabled = True
                Unload Me
                Set frmClientesProveedoresNuevo = Nothing
                Exit Sub
            
            Else
            
                Unload frmClientesProveedoresNuevo
                Set frmClientesProveedoresNuevo = Nothing
                frmClientesProveedoresNuevo.Show
                Exit Sub
            
            End If
        
        Else
            
            MsgBox "El nombre ya existe", vbCritical, "Error"
            Text1(0).SetFocus
        
        End If
        
        If Cn.State = 1 Then Cn.Close
    
    Else
        
        MsgBox "El nombre es obligatorio", vbCritical, "Error"
        Text1(0).SetFocus
    
    End If

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    If StTipoClienteProveedor = "Cliente" Then
    
        If InTipoAltaClienteProveedor = 0 Then
            frmMenuInicial.Enabled = True
        End If
        
        If InTipoAltaClienteProveedor = 1 Then
            frmVentasCompras.Enabled = True
        End If
        
    End If
    
    If StTipoClienteProveedor = "Proveedor" Then
        frmMenuInicial.Enabled = True
    End If
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Sub
