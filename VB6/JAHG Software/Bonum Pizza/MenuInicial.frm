VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenuInicial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Inicial"
   ClientHeight    =   3660
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   9240
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
   Icon            =   "MenuInicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   3240
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
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
      Picture         =   "MenuInicial.frx":10CA
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
      Begin VB.Menu Bebidas 
         Caption         =   "Barra"
         Begin VB.Menu BebidasNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^F
         End
         Begin VB.Menu BebidasExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu Cocina 
         Caption         =   "Cocina"
         Begin VB.Menu CocinaNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^H
         End
         Begin VB.Menu CocinaExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu Ingredientes 
         Caption         =   "Ingredientes"
         Begin VB.Menu IngredientesCombinadosNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^K
         End
         Begin VB.Menu IngredientesCombinadosExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu ListasDeIngredientes 
         Caption         =   "Listas de Ingredientes"
         Begin VB.Menu ListasDeIngredientesBebidas 
            Caption         =   "Barra"
            Begin VB.Menu ListasDeIngredientesBebidasNuevo 
               Caption         =   "Nuevo"
               Shortcut        =   ^L
            End
            Begin VB.Menu ListasDeIngredientesBebidasExistente 
               Caption         =   "Existente"
            End
         End
         Begin VB.Menu ListasDeIngredientesCocina 
            Caption         =   "Cocina"
            Begin VB.Menu ListasDeIngredientesCocinaNuevo 
               Caption         =   "Nuevo"
               Shortcut        =   ^M
            End
            Begin VB.Menu ListasDeIngredientesCocinaExistente 
               Caption         =   "Existente"
            End
         End
      End
      Begin VB.Menu OtrosArticulos 
         Caption         =   "Otros Artículos"
         Begin VB.Menu OtrosArticulosNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^N
         End
         Begin VB.Menu OtrosArticulosExistente 
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
         Begin VB.Menu VentasBarra 
            Caption         =   "Barra"
            Begin VB.Menu VentasBarraLocal 
               Caption         =   "Local"
            End
            Begin VB.Menu VentasBarraDomicilio 
               Caption         =   "Domicilio"
            End
         End
         Begin VB.Menu VentasPagadasCocina 
            Caption         =   "Cocina"
            Begin VB.Menu VentasCocinaLocal 
               Caption         =   "Local"
            End
            Begin VB.Menu VentasCocinaDomicilio 
               Caption         =   "Domicilio"
            End
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
      Begin VB.Menu SalidaInsumos 
         Caption         =   "Salida de Insumos"
      End
      Begin VB.Menu Stock 
         Caption         =   "Stock"
         Begin VB.Menu StockBarra 
            Caption         =   "Barra"
         End
         Begin VB.Menu StockCocina 
            Caption         =   "Cocina"
         End
         Begin VB.Menu StockOtros 
            Caption         =   "Otros"
         End
      End
      Begin VB.Menu MovimientosDeInventario 
         Caption         =   "Movimientos de Inventario"
         Begin VB.Menu MovimientosDeInventarioBarra 
            Caption         =   "Barra"
         End
         Begin VB.Menu MovimientosDeInventarioCocina 
            Caption         =   "Cocina"
         End
         Begin VB.Menu MovimientosDeInventarioOtros 
            Caption         =   "Otros"
         End
      End
      Begin VB.Menu AjustesDeInventario 
         Caption         =   "Ajustes de Inventario"
         Begin VB.Menu AjustesDeInventarioBarra 
            Caption         =   "Barra"
            Shortcut        =   ^T
         End
         Begin VB.Menu AjustesDeInventarioCocina 
            Caption         =   "Cocina"
            Shortcut        =   ^U
         End
         Begin VB.Menu AjustesDeInventarioOtros 
            Caption         =   "Otros"
            Shortcut        =   ^V
         End
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
Private Sub Form_Resize()

    On Error Resume Next
    
    Text1.Top = frmMenuInicial.Height - 1500

End Sub

Private Sub Timer1_Timer()

    On Error Resume Next
    
    Text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
    
End Sub

Private Sub Respaldar_Click()

    On Error GoTo err
    
    FileCopy App.Path & "\DataBase.db", App.Path & "\Respaldos\DataBase" & Replace(Date, "/", "") & ".bck"
    
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
            FileCopy .FileName, App.Path & "\DataBase.db"
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

Private Sub BebidasNuevo_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Barra"
    StSerieItem = "B"
    
    With frmItemNuevo
        .Show
        .Caption = "Añadir nuevo artículo de la Barra"
    End With

End Sub

Private Sub BebidasExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Barra"
    frmItemExistente.Show
    
End Sub

Private Sub CocinaNuevo_Click()

    On Error Resume Next
    
    Me.Enabled = False
    StTipoItem = "Cocina"
    StSerieItem = "C"
    
    With frmItemNuevo
        .Show
        .Caption = "Añadir nuevo artículo de la Cocina"
    End With
    
End Sub

Private Sub CocinaExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Cocina"
    frmItemExistente.Show
    
End Sub

Private Sub IngredientesCombinadosNuevo_Click()
    
    On Error Resume Next
    
    Me.Enabled = False
    StTipoItem = "Ingredientes Generales"
    StSerieItem = "IG"
    
    With frmItemNuevo
        .Show
        .Caption = "Añadir nuevo ingrediente general"
    End With
    
End Sub

Private Sub IngredientesCombinadosExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Ingredientes Generales"
    frmItemExistente.Show
    
End Sub

Private Sub ListasDeIngredientesBebidasNuevo_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Barra"
    
    With frmListaIngredientesNuevo
        .Show
        .Caption = "Nueva lista de ingredientes Barra"
    End With

End Sub

Private Sub ListasDeIngredientesBebidasExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Barra"
    
    With frmListaIngredientesExistente
        .Show
        .Caption = "Listas de ingredientes Barra"
    End With

End Sub

Private Sub ListasDeIngredientesCocinaNuevo_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Cocina"
    
    With frmListaIngredientesNuevo
        .Show
        .Caption = "Nueva lista de ingredientes Cocina"
    End With
    
End Sub

Private Sub ListasDeIngredientesCocinaExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Cocina"
    
    With frmListaIngredientesExistente
        .Show
        .Caption = "Listas de ingredientes Cocina"
    End With
    
End Sub

Private Sub OtrosArticulosNuevo_Click()

    On Error Resume Next
    
    Me.Enabled = False
    StTipoItem = "Otros"
    StSerieItem = "O"
    
    With frmItemNuevo
        .Show
        .Caption = "Añadir nuevo artículo de Otros"
    End With
    
End Sub

Private Sub OtrosArticulosExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    StTipoItem = "Otros"
    
    frmItemExistente.Show
    
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

Private Sub VentasBarraLocal_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Barra Local"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Historial de ventas barra local"
        .Pagar.Visible = False
        .Agregar.Visible = False
    End With
    
End Sub

Private Sub VentasBarraDomicilio_Click()
    
    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Barra Domicilio"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Historial de ventas barra domicilio"
        .Pagar.Visible = False
        .Agregar.Visible = False
    End With
    
End Sub

Private Sub VentasCocinaLocal_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Cocina Local"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Historial de ventas cocina local"
        .Pagar.Visible = False
        .Agregar.Visible = False
    End With
    
End Sub

Private Sub VentasCocinaDomicilio_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoVentasCompras = "Ventas"
    
    StTipoVenta = "Cocina Domicilio"
    
    With frmHistorialVentasCompras
        .Show
        .Caption = "Historial de ventas cocina domicilio"
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

    On Error Resume Next

    Me.Enabled = False
    
    StTipoClienteProveedor = "Proveedor"
    
    With frmClientesProveedoresNuevo
        .Show
        .Caption = "Añadir nuevo Proveedor"
    End With
    
End Sub

Private Sub ProveedoresExistente_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoClienteProveedor = "Proveedor"
    
    With frmClientesProveedoresExistente
        .Show
        .Caption = "Catalogo de Proveedores"
    End With
    
End Sub

Private Sub ComprasPagadas_Click()

    On Error Resume Next

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

    On Error Resume Next

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

Private Sub SalidaInsumos_Click()

    On Error Resume Next
    
    Me.Enabled = False
    
    frmSalidaInsumos.Show
    
End Sub

Private Sub StockBarra_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Barra"
    
    frmStock.Show
    
End Sub

Private Sub StockCocina_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Cocina"
    
    frmStock.Show
    
End Sub

Private Sub StockOtros_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Otros"
    
    frmStock.Show
    
End Sub

Private Sub MovimientosDeInventarioBarra_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Barra"
    
    frmMovimientosInventario.Show
    
End Sub

Private Sub MovimientosDeInventarioCocina_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Cocina"
    
    frmMovimientosInventario.Show
    
End Sub

Private Sub MovimientosDeInventarioOtros_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Otros"
    
    frmMovimientosInventario.Show
    
End Sub

Private Sub AjustesDeInventarioBarra_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Barra"
    
    frmAjusteInventario.Show
    
End Sub

Private Sub AjustesDeInventarioCocina_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Cocina"
    
    frmAjusteInventario.Show
    
End Sub

Private Sub AjustesDeInventarioOtros_Click()

    On Error Resume Next

    Me.Enabled = False
    
    StTipoItem = "Otros"
    
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
    
    Unload TicketComprasVentas
    Unload TicketCorte
    Unload TicketPedidos
    
    Unload Me
    
End Sub





