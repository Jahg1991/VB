VERSION 5.00
Begin VB.Form frmInicioSesion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2760
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6975
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
   ScaleHeight     =   2760
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   1
      Left            =   5040
      Picture         =   "InicioSesion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   0
      Left            =   480
      Picture         =   "InicioSesion.frx":068B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6495
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
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
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
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmInicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim In1 As Integer

Private Sub Command1_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            With Cn
                .CursorLocation = adUseClient
                .Open StConnection
            End With
            
            With Rs
                If .State = 1 Then .Close
                .Open "Select count(*) as existe from Usuarios where nombre like '" & Text1(0) & "' and contrasena like '" & Text1(1) & "';", Cn, adOpenStatic, adLockOptimistic
                .Requery
                In1 = .Fields(0).Value
                .Close
            End With
            
            If In1 = 1 Then
                
                With Rs
                    If .State = 1 Then .Close
                    .Open "Select * from Usuarios where nombre like '" & Text1(0) & "' and contrasena like '" & Text1(1) & "';", Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    StUsuario = .Fields(1).Value
                    StPermisosArchivo = .Fields(3).Value
                    StPermisosArticulos = .Fields(4).Value
                    StPermisosVentas = .Fields(5).Value
                    StPermisosCompras = .Fields(6).Value
                    StPermisosInventario = .Fields(7).Value
                    StPermisosCorteCaja = .Fields(8).Value
                    StPermisosProduccion = .Fields(9).Value
                    .Close
                End With
                
                With frmMenuInicial
                    
                    .Show
                    .Caption = StUsuario
                    
                    If StPermisosArchivo = "Si" Then
                        .Archivo.Visible = True
                    Else
                        .Archivo.Visible = False
                    End If
                    
                    If StPermisosArticulos = "Si" Then
                        .Articulos.Visible = True
                    Else
                        .Articulos.Visible = False
                    End If
                    
                    If StPermisosVentas = "Si" Then
                        .Ventas.Visible = True
                    Else
                        .Ventas.Visible = False
                    End If
                    
                    If StPermisosCompras = "Si" Then
                        .Compras.Visible = True
                    Else
                        .Compras.Visible = False
                    End If
                    
                    If StPermisosInventario = "Si" Then
                        .Inventario.Visible = True
                    Else
                        .Inventario.Visible = False
                    End If
                    
                    If StPermisosCorteCaja = "Si" Then
                        .CorteDeCaja.Visible = True
                    Else
                        frmMenuInicial.CorteDeCaja.Visible = False
                    End If
                    
                    If StPermisosProduccion = "Si" Then
                        .Produccion.Visible = True
                    Else
                        frmMenuInicial.Produccion.Visible = False
                    End If
                    
                    With .Image1
                        .Width = frmMenuInicial.Width
                        .Height = frmMenuInicial.Height
                    End With
                
                End With
                
                Unload Me
            
            Else
                
                MsgBox "Usuario o contraseña incorrectos", vbCritical, "Error"
                Unload frmInicioSesion
                Set frmInicioSesion = Nothing
                frmInicioSesion.Show
            
            End If
        
        Case 1
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Sub
