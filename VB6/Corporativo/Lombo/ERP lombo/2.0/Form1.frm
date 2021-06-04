VERSION 5.00
Begin VB.Form MenuInicial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pizzeria"
   ClientHeight    =   3135
   ClientLeft      =   150
   ClientTop       =   1095
   ClientWidth     =   4680
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
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Respaldar 
         Caption         =   "Respaldar"
      End
      Begin VB.Menu Restaurar 
         Caption         =   "Restaurar"
      End
   End
   Begin VB.Menu Articulos 
      Caption         =   "Artículos"
      Begin VB.Menu Bebidas 
         Caption         =   "Bebidas"
         Begin VB.Menu BebidasNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu BebidasExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu Ingredientes 
         Caption         =   "Ingredientes"
         Begin VB.Menu IngredientesNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu IngredientesExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu Pizzas 
         Caption         =   "Pizzas"
         Begin VB.Menu PizzasNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu PizzasExistente 
            Caption         =   "Exixtente"
         End
      End
      Begin VB.Menu ListasDeIngredientes 
         Caption         =   "Listas de Ingredientes"
         Begin VB.Menu ListasDeIngredientesNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu ListaDeIngredientesExistente 
            Caption         =   "Existente"
         End
      End
   End
   Begin VB.Menu Ventas 
      Caption         =   "Ventas"
      Index           =   1
      Begin VB.Menu Clientes 
         Caption         =   "Clientes"
         Begin VB.Menu ClientesNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu ClientesExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu HistorialDeVentas 
         Caption         =   "Historial de Ventas"
         Begin VB.Menu VentasPagadas 
            Caption         =   "Pagadas"
            Begin VB.Menu VentasPagadasLocal 
               Caption         =   "Local"
            End
            Begin VB.Menu VentasPagadasADomicilio 
               Caption         =   "A Domicilio"
            End
         End
         Begin VB.Menu VentasNoPagadas 
            Caption         =   "No Pagadas"
            Begin VB.Menu VentasNoPagadasLocal 
               Caption         =   "Local"
            End
            Begin VB.Menu VentasNoPagadasADomicilio 
               Caption         =   "A Domicilio"
            End
         End
      End
      Begin VB.Menu VentasVentas 
         Caption         =   "Ventas"
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
      Begin VB.Menu Proveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu ProveedoresNuevo 
            Caption         =   "Nuevo"
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
      Begin VB.Menu ComprasCompras 
         Caption         =   "Compras"
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
      End
      Begin VB.Menu SalidaDeDinero 
         Caption         =   "Salida de Dinero"
      End
      Begin VB.Menu CorteDeCajaCorteDeCaja 
         Caption         =   "Corte de Caja"
      End
   End
   Begin VB.Menu AcercaDe 
      Caption         =   "Acerca de"
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MenuInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Salir_Click()
    Unload Me
End Sub
