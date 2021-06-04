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
      BackColor       =   &H0000C000&
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
    Option Explicit
    
    '//RECORDSET
    Dim Rs      As New adodb.Recordset
    
    '//OTROS
    Dim In1     As Long
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                With Cn
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    If .State = 0 Then .Open (StConnection)
                End With
                
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from FND_USERS where nombre like '" & Text1(0).Text & "' and contrasena like '" & Text1(1).Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    
                    In1 = .Fields(0).Value
                    
                    .Close
                End With
                
                If In1 = 1 Then
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from FND_USERS where nombre like '" & Text1(0).Text & "' and contrasena like '" & Text1(1).Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        
                        StUsuario = .Fields(1).Value
                        If IsNull(.Fields(3).Value) = False Then StPermisosArchivo = .Fields(3).Value Else StPermisosArchivo = "No"
                        If IsNull(.Fields(4).Value) = False Then StPermisosArticulos = .Fields(4).Value Else StPermisosArticulos = "No"
                        If IsNull(.Fields(5).Value) = False Then StPermisosVentas = .Fields(5).Value Else StPermisosVentas = "No"
                        If IsNull(.Fields(6).Value) = False Then StPermisosCompras = .Fields(6).Value Else StPermisosCompras = "No"
                        If IsNull(.Fields(7).Value) = False Then StPermisosInventario = .Fields(7).Value Else StPermisosInventario = "No"
                        If IsNull(.Fields(8).Value) = False Then StPermisosCorteCaja = .Fields(8).Value Else StPermisosCorteCaja = "No"
                        If IsNull(.Fields(9).Value) = False Then StPermisosProduccion = .Fields(9).Value Else StPermisosProduccion = "No"
                        If IsNull(.Fields(10).Value) = False Then StPermisosCaja = .Fields(10).Value Else StPermisosCaja = "No"
                        
                        .Close
                    End With
                    
                    With frmMenuInicial
                        .Show
                        
                        .Caption = "Punto de venta " & PcNombreEmpresa & ", Usuario activo: " & StUsuario
                        
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
                            .CorteDeCaja.Visible = False
                        End If
                        
                        If StPermisosProduccion = "Si" Then
                            .Produccion.Visible = True
                        Else
                            .Produccion.Visible = False
                        End If
                        
                        With .Image1
                            .Width = frmMenuInicial.Width
                            .Height = frmMenuInicial.Height
                        End With
                        
                        With .Text1
                            .Text = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
                            .Top = frmMenuInicial.Height - 1500
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
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmInicioSesion:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmInicioSesion:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
