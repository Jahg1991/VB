VERSION 5.00
Begin VB.Form frmInicioSesion 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9825
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17475
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
   ScaleHeight     =   9825
   ScaleWidth      =   17475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   7680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   0
      Left            =   7680
      TabIndex        =   0
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   5370
      TabIndex        =   4
      Top             =   3645
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6495
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "SALIR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ACEPTAR"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CONTRASEÑA"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "USUARIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   495
            Index           =   0
            Left            =   720
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
'***********************************************************************************
'Nombre:        frmInicioSesion
'Proposito:     Validar inicio de sesion y establecer permisos
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    '//RECORDSET
    Dim Rs      As New adodb.Recordset
    '//OTROS
    Dim In1     As Long
    Dim i       As Integer
    
    Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                If KeyAscii = 13 Then
                    With Text1(1)
                        .SetFocus
                    End With
                End If
            Case 1
                If KeyAscii = 13 Then
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
                        If In1 = 1 Then
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from FND_USERS where nombre like '" & Text1(0).Text & "' and contrasena like '" & Text1(1).Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            StUsuario = .Fields(1).Value
                            If IsNull(.Fields(3).Value) = False Then StPermisosArchivo = .Fields(3).Value Else StPermisosArchivo = "No"
                            
                            If IsNull(.Fields(4).Value) = False Then StPermisosCatalogos = .Fields(4).Value Else StPermisosCatalogos = "No"
                            
                            If IsNull(.Fields(5).Value) = False Then StPermisosListaMateriales = .Fields(5).Value Else StPermisosListaMateriales = "No"
                            
                            If IsNull(.Fields(6).Value) = False Then StPermisosProduccion = .Fields(6).Value Else StPermisosProduccion = "No"
                            
                            If IsNull(.Fields(7).Value) = False Then StPermisosVentas = .Fields(7).Value Else StPermisosVentas = "No"
                            
                            If IsNull(.Fields(8).Value) = False Then StPermisosPedidos = .Fields(8).Value Else StPermisosVentas = "No"
                            
                            If IsNull(.Fields(9).Value) = False Then StPermisosCompras = .Fields(9).Value Else StPermisosCompras = "No"
                            
                            If IsNull(.Fields(10).Value) = False Then StPermisosAjustes = .Fields(10).Value Else StPermisosAjustes = "No"
                            
                            If IsNull(.Fields(11).Value) = False Then StPermisosInventario = .Fields(11).Value Else StPermisosInventario = "No"
                            
                            If IsNull(.Fields(12).Value) = False Then StPermisosCorteCaja = .Fields(12).Value Else StPermisosCorteCaja = "No"
                            
                            If IsNull(.Fields(13).Value) = False Then StPermisosCaja = .Fields(13).Value Else StPermisosCaja = "No"
                            
                            If IsNull(.Fields(14).Value) = False Then StCajaPredeterminada = .Fields(14).Value Else StCajaPredeterminada = "Caja 10"
                            
                            If IsNull(.Fields(15).Value) = False Then StPermisosRCatalogos = .Fields(4).Value Else StPermisosCatalogos = "No"
                            
                            If IsNull(.Fields(16).Value) = False Then StPermisosRListaMateriales = .Fields(5).Value Else StPermisosListaMateriales = "No"
                            
                            If IsNull(.Fields(17).Value) = False Then StPermisosRProduccion = .Fields(6).Value Else StPermisosProduccion = "No"
                            
                            If IsNull(.Fields(18).Value) = False Then StPermisosRPedidos = .Fields(8).Value Else StPermisosVentas = "No"
                            
                            If IsNull(.Fields(19).Value) = False Then StPermisosRVentas = .Fields(7).Value Else StPermisosVentas = "No"
                            
                            If IsNull(.Fields(20).Value) = False Then StPermisosRCompras = .Fields(9).Value Else StPermisosCompras = "No"
                            
                            If IsNull(.Fields(21).Value) = False Then StPermisosRInventario = .Fields(11).Value Else StPermisosInventario = "No"
                            
                            If IsNull(.Fields(22).Value) = False Then StPermisosRCorteCaja = .Fields(12).Value Else StPermisosCorteCaja = "No"
                            .Close
                            With frmMenuInicial
                                .Show
                                .Caption = "PUNTO DE VENTA " & PcNombreEmpresa & ", USUARIO ACTIVO: " & StUsuario
                                If StPermisosArchivo = "Si" Then
                                    With .Archivo
                                        .Visible = True
                                    End With
                                Else
                                    With .Archivo
                                        .Visible = False
                                    End With
                                End If
                                If StPermisosCatalogos = "Si" Then
                                    With .Catalogos
                                        .Visible = True
                                    End With
                                Else
                                    With .Catalogos
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosListaMateriales = "Si" Then
                                    With .ListasDeIngredientes
                                        .Visible = True
                                    End With
                                Else
                                    With .ListasDeIngredientes
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosProduccion = "Si" Then
                                    With .Produccion
                                        .Visible = True
                                    End With
                                Else
                                    With .Produccion
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosVentas = "Si" Then
                                    With .Ventas
                                        .Visible = True
                                    End With
                                Else
                                    With .Ventas
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosPedidos = "Si" Then
                                    With .Pedidos
                                        .Visible = True
                                    End With
                                Else
                                    With .Pedidos
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosCompras = "Si" Then
                                    With .Compras
                                        .Visible = True
                                    End With
                                Else
                                    With .Compras
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosAjustes = "Si" Then
                                    With .Ajustes
                                        .Visible = True
                                    End With
                                Else
                                    With .Ajustes
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosInventario = "Si" Then
                                    With .Inventario
                                        .Visible = True
                                    End With
                                Else
                                    With .Inventario
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosCorteCaja = "Si" Then
                                    With .CorteDeCaja
                                        .Visible = True
                                    End With
                                Else
                                    With .CorteDeCaja
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosCaja = "Si" Then
                                    With .Combo1
                                        .Enabled = True
                                    End With
                                Else
                                    With .Combo1
                                        .Enabled = False
                                    End With
                                End If
                                
                                If StPermisosRCatalogos = "Si" Then
                                    With .rCatalogos
                                        .Visible = True
                                    End With
                                Else
                                    With .rCatalogos
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRListaMateriales = "Si" Then
                                    With .rListasDeIngredientes
                                        .Visible = True
                                    End With
                                Else
                                    With .rListasDeIngredientes
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRProduccion = "Si" Then
                                    With .rProduccion
                                        .Visible = True
                                    End With
                                Else
                                    With .rProduccion
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRVentas = "Si" Then
                                    With .rVentas
                                        .Visible = True
                                    End With
                                Else
                                    With .rVentas
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRPedidos = "Si" Then
                                    With .rPedidos
                                        .Visible = True
                                    End With
                                Else
                                    With .rPedidos
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRCompras = "Si" Then
                                    With .rCompras
                                        .Visible = True
                                    End With
                                Else
                                    With .rCompras
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRInventario = "Si" Then
                                    With .rInventarios
                                        .Visible = True
                                    End With
                                Else
                                    With .rInventarios
                                        .Visible = False
                                    End With
                                End If
                                
                                If StPermisosRCorteCaja = "Si" Then
                                    With .rCorteDeCaja
                                        .Visible = True
                                    End With
                                Else
                                    With .rCorteDeCaja
                                        .Visible = False
                                    End With
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
                            
                            With frmInicioSesion
                                .Show
                            End With
                        End If
                    End With
                End If
            End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmInicioSesion:Text1_KeyPress" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
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
                    If In1 = 1 Then
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from FND_USERS where nombre like '" & Text1(0).Text & "' and contrasena like '" & Text1(1).Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        StUsuario = .Fields(1).Value
                        If IsNull(.Fields(3).Value) = False Then StPermisosArchivo = .Fields(3).Value Else StPermisosArchivo = "No"
                        
                        If IsNull(.Fields(4).Value) = False Then StPermisosCatalogos = .Fields(4).Value Else StPermisosCatalogos = "No"
                        
                        If IsNull(.Fields(5).Value) = False Then StPermisosListaMateriales = .Fields(5).Value Else StPermisosListaMateriales = "No"
                        
                        If IsNull(.Fields(6).Value) = False Then StPermisosProduccion = .Fields(6).Value Else StPermisosProduccion = "No"
                        
                        If IsNull(.Fields(7).Value) = False Then StPermisosVentas = .Fields(7).Value Else StPermisosVentas = "No"
                        
                        If IsNull(.Fields(8).Value) = False Then StPermisosPedidos = .Fields(8).Value Else StPermisosVentas = "No"
                        
                        If IsNull(.Fields(9).Value) = False Then StPermisosCompras = .Fields(9).Value Else StPermisosCompras = "No"
                        
                        If IsNull(.Fields(10).Value) = False Then StPermisosAjustes = .Fields(10).Value Else StPermisosAjustes = "No"
                        
                        If IsNull(.Fields(11).Value) = False Then StPermisosInventario = .Fields(11).Value Else StPermisosInventario = "No"
                        
                        If IsNull(.Fields(12).Value) = False Then StPermisosCorteCaja = .Fields(12).Value Else StPermisosCorteCaja = "No"
                        
                        If IsNull(.Fields(13).Value) = False Then StPermisosCaja = .Fields(13).Value Else StPermisosCaja = "No"
                        
                        If IsNull(.Fields(14).Value) = False Then StCajaPredeterminada = .Fields(14).Value Else StCajaPredeterminada = "Caja 10"
                        
                        If IsNull(.Fields(15).Value) = False Then StPermisosRCatalogos = .Fields(4).Value Else StPermisosCatalogos = "No"
                        
                        If IsNull(.Fields(16).Value) = False Then StPermisosRListaMateriales = .Fields(5).Value Else StPermisosListaMateriales = "No"
                        
                        If IsNull(.Fields(17).Value) = False Then StPermisosRProduccion = .Fields(6).Value Else StPermisosProduccion = "No"
                        
                        If IsNull(.Fields(18).Value) = False Then StPermisosRPedidos = .Fields(8).Value Else StPermisosVentas = "No"
                        
                        If IsNull(.Fields(19).Value) = False Then StPermisosRVentas = .Fields(7).Value Else StPermisosVentas = "No"
                        
                        If IsNull(.Fields(20).Value) = False Then StPermisosRCompras = .Fields(9).Value Else StPermisosCompras = "No"
                        
                        If IsNull(.Fields(21).Value) = False Then StPermisosRInventario = .Fields(11).Value Else StPermisosInventario = "No"
                        
                        If IsNull(.Fields(22).Value) = False Then StPermisosRCorteCaja = .Fields(12).Value Else StPermisosCorteCaja = "No"
                        .Close
                        With frmMenuInicial
                            .Show
                            .Caption = "PUNTO DE VENTA " & PcNombreEmpresa & ", USUARIO ACTIVO: " & StUsuario
                            If StPermisosArchivo = "Si" Then
                                With .Archivo
                                    .Visible = True
                                End With
                            Else
                                With .Archivo
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosCatalogos = "Si" Then
                                With .Catalogos
                                    .Visible = True
                                End With
                            Else
                                With .Catalogos
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosListaMateriales = "Si" Then
                                With .ListasDeIngredientes
                                    .Visible = True
                                End With
                            Else
                                With .ListasDeIngredientes
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosProduccion = "Si" Then
                                With .Produccion
                                    .Visible = True
                                End With
                            Else
                                With .Produccion
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosVentas = "Si" Then
                                With .Ventas
                                    .Visible = True
                                End With
                            Else
                                With .Ventas
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosPedidos = "Si" Then
                                With .Pedidos
                                    .Visible = True
                                End With
                            Else
                                With .Pedidos
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosCompras = "Si" Then
                                With .Compras
                                    .Visible = True
                                End With
                            Else
                                With .Compras
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosAjustes = "Si" Then
                                With .Ajustes
                                    .Visible = True
                                End With
                            Else
                                With .Ajustes
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosInventario = "Si" Then
                                With .Inventario
                                    .Visible = True
                                End With
                            Else
                                With .Inventario
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosCorteCaja = "Si" Then
                                With .CorteDeCaja
                                    .Visible = True
                                End With
                            Else
                                With .CorteDeCaja
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosCaja = "Si" Then
                                With .Combo1
                                    .Enabled = True
                                End With
                            Else
                                With .Combo1
                                    .Enabled = False
                                End With
                            End If
                            
                            If StPermisosRCatalogos = "Si" Then
                                With .rCatalogos
                                    .Visible = True
                                End With
                            Else
                                With .rCatalogos
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRListaMateriales = "Si" Then
                                With .rListasDeIngredientes
                                    .Visible = True
                                End With
                            Else
                                With .rListasDeIngredientes
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRProduccion = "Si" Then
                                With .rProduccion
                                    .Visible = True
                                End With
                            Else
                                With .rProduccion
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRVentas = "Si" Then
                                With .rVentas
                                    .Visible = True
                                End With
                            Else
                                With .rVentas
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRPedidos = "Si" Then
                                With .rPedidos
                                    .Visible = True
                                End With
                            Else
                                With .rPedidos
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRCompras = "Si" Then
                                With .rCompras
                                    .Visible = True
                                End With
                            Else
                                With .rCompras
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRInventario = "Si" Then
                                With .rInventarios
                                    .Visible = True
                                End With
                            Else
                                With .rInventarios
                                    .Visible = False
                                End With
                            End If
                            
                            If StPermisosRCorteCaja = "Si" Then
                                With .rCorteDeCaja
                                    .Visible = True
                                End With
                            Else
                                With .rCorteDeCaja
                                    .Visible = False
                                End With
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
                        
                        With frmInicioSesion
                            .Show
                        End With
                    End If
                End With
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
        With Rs
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
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
