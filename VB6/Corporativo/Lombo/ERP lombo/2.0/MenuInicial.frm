VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenuInicial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Inicial"
   ClientHeight    =   3525
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   8985
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
   ScaleHeight     =   3525
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   3555
      Left            =   0
      Picture         =   "MenuInicial.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Respaldar 
         Caption         =   "Respaldar"
         Shortcut        =   ^A
      End
      Begin VB.Menu Restaurar 
         Caption         =   "Restaurar"
         Shortcut        =   ^B
      End
      Begin VB.Menu Usuarios 
         Caption         =   "Usuarios"
         Begin VB.Menu UsuariosNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^C
         End
         Begin VB.Menu UsuariosExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu Opciones 
         Caption         =   "Opciones"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu Articulos 
      Caption         =   "Artículos"
      Begin VB.Menu ItemsNuevo 
         Caption         =   "Nuevo"
         Shortcut        =   ^F
      End
      Begin VB.Menu ItemsExistente 
         Caption         =   "Existente"
      End
   End
   Begin VB.Menu Produccion 
      Caption         =   "Produccion"
      Begin VB.Menu ListasDeIngredientes 
         Caption         =   "Listas de Ingredientes"
         Begin VB.Menu ListasDeIngredientesNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^L
         End
         Begin VB.Menu ListasDeIngredientesExistente 
            Caption         =   "Existente"
         End
      End
   End
   Begin VB.Menu Ventas 
      Caption         =   "Ventas"
      Begin VB.Menu Clientes 
         Caption         =   "Clientes"
         Begin VB.Menu ClientesNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^O
         End
         Begin VB.Menu ClientesExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu HistorialDeVentas 
         Caption         =   "Historial de Ventas"
         Begin VB.Menu VentasLocal 
            Caption         =   "Local"
         End
         Begin VB.Menu VentasDomicilio 
            Caption         =   "Domicilio"
         End
      End
      Begin VB.Menu VentasNoPagadas 
         Caption         =   "Cuenta abierta"
         Shortcut        =   ^P
      End
      Begin VB.Menu NuevaVenta 
         Caption         =   "Nueva Venta"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
      Begin VB.Menu Proveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu ProveedoresNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^R
         End
         Begin VB.Menu ProveedoresExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu HistorialDeCompras 
         Caption         =   "Historial de Compras"
         Begin VB.Menu ComprasPagadas 
            Caption         =   "Pagadas"
         End
         Begin VB.Menu ComprasNoPagadas 
            Caption         =   "No Pagadas"
         End
      End
      Begin VB.Menu NuevaCompra 
         Caption         =   "Nueva Compra"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Inventario 
      Caption         =   "Inventario"
      Begin VB.Menu Stock 
         Caption         =   "Stock"
      End
      Begin VB.Menu MovimientosDeInventario 
         Caption         =   "Movimientos de Inventario"
      End
      Begin VB.Menu AjustesDeInventario 
         Caption         =   "Ajustes de Inventario"
      End
   End
   Begin VB.Menu CorteDeCaja 
      Caption         =   "Corte de Caja"
      Begin VB.Menu EntradaDeDinero 
         Caption         =   "Entarda de Dinero"
         Shortcut        =   ^W
      End
      Begin VB.Menu SalidaDeDinero 
         Caption         =   "Salida de Dinero"
         Shortcut        =   ^X
      End
      Begin VB.Menu CorteDeCajaCorteDeCaja 
         Caption         =   "Corte de Caja"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu AcercaDe 
      Caption         =   "Acerca de"
   End
   Begin VB.Menu CerrarSesion 
      Caption         =   "Cerrar Sesion"
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmMenuInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Respaldar_Click()

    On Error GoTo err
    
    FileCopy App.Path & "\DataBase.mdb", App.Path & "\Respaldos\DataBase" & Replace(Date, "/", "") & ".bck"
    
    MsgBox "Respaldo realizado con éxito", vbOKOnly, "Terminado"
    
    Exit Sub
    
err:
    MsgBox "Error al realizar el respaldo - " & err.Description, vbCritical, "Error"
    
End Sub

Private Sub Restaurar_Click()

    On Error GoTo err
    
    With CommonDialog1
        .DialogTitle = "Seleccione el respaldo"
        .Filter = "Archivos de respaldo|*.bck"
        .InitDir = App.Path & "\Respaldos"
        .ShowOpen
    
        If .FileName <> "" Then
            FileCopy .FileName, App.Path & "\DataBase.mdb"
            MsgBox "Restauración realizada con éxito", vbOKOnly, "Terminado"
        End If
        
    End With
    
    Exit Sub
    
err:
    MsgBox "Error al realizar la restauración - " & err.Description, vbCritical, "Error"
    
End Sub

Private Sub UsuariosNuevo_Click()

    On Error Resume Next
    
    frmUsuariosNuevo.Show
    Me.Enabled = False
    
End Sub

Private Sub UsuariosExistente_Click()

    On Error Resume Next

    frmUsuariosExistente.Show
    Me.Enabled = False
    
End Sub

Private Sub Opciones_Click()

    On Error Resume Next
    
    Me.Enabled = False
    frmOpciones.Show
    
End Sub

Private Sub ItemsNuevo_Click()

    On Error Resume Next

    Me.Enabled = False
    
    With frmItemNuevo
        .Show
        .Caption = "Añadir nuevo artículo"
    End With

End Sub

Private Sub ItemsExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    frmItemExistente.Show
    
End Sub

Private Sub ListasDeIngredientesNuevo_Click()

    On Error Resume Next

    Me.Enabled = False
    
    With frmListaIngredientesNuevo
        .Show
        .Caption = "Nueva lista de ingredientes"
    End With
    
End Sub

Private Sub ListasDeIngredientesExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    
    With frmListaIngredientesExistente
        .Show
        .Caption = "Listas de ingredientes"
    End With
    
End Sub

Private Sub ClientesNuevo_Click()
    
    On Error Resume Next
    
    InTipoAltaClienteProveedor = 0
    Me.Enabled = False
    StTipoClienteProveedor = "Cliente"
    
    With frmClientesProveedoresNuevo
        .Show
        .Caption = "Añadir nuevo Cliente"
    End With

End Sub

Private Sub ClientesExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoClienteProveedor = "Cliente"
    
    With frmClientesProveedoresExistente
        .Show
        .Caption = "Catálogo de Clientes"
    End With
    
End Sub

Private Sub VentasLocal_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Local"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Historial de ventas local"
        .Pagar.Visible = False
        .Agregar.Visible = False
    End With
    
End Sub

Private Sub VentasDomicilio_Click()
    
    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Domicilio"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Historial de ventas a domicilio"
        .Pagar.Visible = False
        .Agregar.Visible = False
    End With
    
End Sub

Private Sub VentasNoPagadas_Click()
    
    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Abiertas"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Cuenta Abierta"
        .Pagar.Visible = True
        .Agregar.Visible = True
    End With
    
End Sub

Private Sub NuevaVenta_Click()
    
    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    With frmVentasCompras
        .Show
        .Caption = "Nueva Venta"
    End With
    
End Sub

Private Sub ProveedoresNuevo_Click()

    Me.Enabled = False
    StTipoClienteProveedor = "Proveedor"
    
    With frmClientesProveedoresNuevo
        .Show
        .Caption = "Añadir nuevo Proveedor"
    End With
    
End Sub

Private Sub ProveedoresExistente_Click()

    Me.Enabled = False
    
    StTipoClienteProveedor = "Proveedor"
    
    With frmClientesProveedoresExistente
        .Show
        .Caption = "Catalogo de Proveedores"
    End With
    
End Sub

Private Sub ComprasPagadas_Click()

    Me.Enabled = False
    
    StTipoVentasCompras = "Compras"
    
    StTipoCompra = "Pagadas"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Compras Pagadas"
        .Pagar.Visible = False
        .Agregar.Visible = False
    End With
    
End Sub

Private Sub ComprasNoPagadas_Click()

    Me.Enabled = False
    
    StTipoVentasCompras = "Compras"
    
    StTipoCompra = "No Pagadas"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Compras no Pagadas"
        .Pagar.Visible = True
        .Agregar.Visible = True
    End With
    
End Sub

Private Sub NuevaCompra_Click()
    
    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Compras"
    
    With frmVentasCompras
        .Show
        .Caption = "Nueva Compra"
    End With
    
End Sub

Private Sub Stock_Click()
    
    On Error Resume Next

    Me.Enabled = False
    frmStock.Show
    
End Sub

Private Sub MovimientosDeInventario_Click()

    On Error Resume Next

    Me.Enabled = False
    frmMovimientosInventario.Show
    
End Sub

Private Sub AjustesDeInventario_Click()

    On Error Resume Next

    Me.Enabled = False
    frmAjusteInventario.Show
    
End Sub

Private Sub EntradaDeDinero_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoEntradaSalida = "Entrada"
    
    With frmEntradaSalidaDinero
        .Show
        .Caption = "Entrada de efectivo"
    End With
    
End Sub

Private Sub SalidaDeDinero_Click()

    On Error Resume Next
    
    Me.Enabled = False
    StTipoEntradaSalida = "Salida"
    
    With frmEntradaSalidaDinero
        .Show
        .Caption = "Salida de efectivo"
    End With
    
End Sub

Private Sub CorteDeCajaCorteDeCaja_Click()

    On Error Resume Next

    Me.Enabled = False
    frmCorteCaja.Show
    
End Sub

Private Sub AcercaDe_Click()

    On Error Resume Next

    Me.Enabled = False
    frmAcercaDe.Show
    
End Sub

Private Sub CerrarSesion_Click()
    
    On Error Resume Next

    Unload Me
    Main

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next

    Unload Me
    
End Sub





