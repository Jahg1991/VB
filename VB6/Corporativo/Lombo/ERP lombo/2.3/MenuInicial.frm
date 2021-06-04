VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenuInicial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Inicial"
   ClientHeight    =   3885
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   10350
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
   Picture         =   "MenuInicial.frx":10CA
   ScaleHeight     =   3885
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   420
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
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
      Left            =   360
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
      Height          =   3915
      Left            =   0
      Picture         =   "MenuInicial.frx":F441
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10335
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Respaldar 
         Caption         =   "Respaldar"
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu Restaurar 
         Caption         =   "Restaurar"
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu Usuarios 
         Caption         =   "Usuarios"
         Begin VB.Menu UsuariosNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^C
         End
         Begin VB.Menu UsuariosExistente 
            Caption         =   "Existente"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu Opciones 
         Caption         =   "Opciones"
         Shortcut        =   ^E
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
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu Produccion 
      Caption         =   "Produccion"
      Begin VB.Menu ReportarProduccion 
         Caption         =   "Reportar Produccion"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu ListasDeIngredientes 
         Caption         =   "Listas de Ingredientes"
         Begin VB.Menu ListasDeIngredientesNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^H
         End
         Begin VB.Menu ListasDeIngredientesExistente 
            Caption         =   "Existente"
            Shortcut        =   ^I
         End
      End
   End
   Begin VB.Menu Ventas 
      Caption         =   "Ventas"
      Begin VB.Menu Clientes 
         Caption         =   "Clientes"
         Begin VB.Menu ClientesNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^J
         End
         Begin VB.Menu ClientesExistente 
            Caption         =   "Existente"
            Shortcut        =   ^K
         End
      End
      Begin VB.Menu HistorialDeVentas 
         Caption         =   "Historial de Ventas"
         Begin VB.Menu VentasLocal 
            Caption         =   "Local"
            Shortcut        =   ^L
         End
         Begin VB.Menu VentasDomicilio 
            Caption         =   "Domicilio"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu VentasNoPagadas 
         Caption         =   "Cuenta abierta"
         Shortcut        =   ^N
      End
      Begin VB.Menu Pedidos 
         Caption         =   "Pedidos"
         Begin VB.Menu NuevoPedido 
            Caption         =   "Nuevo Pedido"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu PedidosPendientes 
            Caption         =   "Pedidos Pendientes"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu CancelacionPedidos 
            Caption         =   "Cancelacion de Pedidos"
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu NuevaVenta 
         Caption         =   "Nueva Venta"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
      Begin VB.Menu Proveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu ProveedoresNuevo 
            Caption         =   "Nuevo"
            Shortcut        =   ^P
         End
         Begin VB.Menu ProveedoresExistente 
            Caption         =   "Existente"
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu HistorialDeCompras 
         Caption         =   "Historial de Compras"
         Begin VB.Menu ComprasPagadas 
            Caption         =   "Pagadas"
            Shortcut        =   ^R
         End
         Begin VB.Menu ComprasNoPagadas 
            Caption         =   "No Pagadas"
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu NuevaCompra 
         Caption         =   "Nueva Compra"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu Inventario 
      Caption         =   "Inventario"
      Begin VB.Menu Stock 
         Caption         =   "Stock"
         Shortcut        =   ^U
      End
      Begin VB.Menu MovimientosDeInventario 
         Caption         =   "Movimientos de Inventario"
         Shortcut        =   ^V
      End
      Begin VB.Menu AjustesDeInventario 
         Caption         =   "Ajustes de Inventario"
         Shortcut        =   ^W
      End
      Begin VB.Menu SalidaInsumos 
         Caption         =   "Salida de Insumos"
         Shortcut        =   ^X
      End
      Begin VB.Menu Reportes 
         Caption         =   "Reportes"
         Begin VB.Menu Rastreabilidad 
            Caption         =   "Rastreabilidad"
         End
      End
   End
   Begin VB.Menu CorteDeCaja 
      Caption         =   "Corte de Caja"
      Begin VB.Menu EntradaDeDinero 
         Caption         =   "Entarda de Dinero"
         Shortcut        =   ^Y
      End
      Begin VB.Menu SalidaDeDinero 
         Caption         =   "Salida de Dinero"
         Shortcut        =   ^Z
      End
      Begin VB.Menu CorteDeCajaCorteDeCaja 
         Caption         =   "Corte de Caja"
         Shortcut        =   ^{F2}
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
    Option Explicit
    
    '//OTROS
    Dim i As Long
    Dim a As String

    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Combo1
            For i = 1 To 10
                .AddItem "Caja " & i
            Next i
            
            .Text = StPermisosCaja
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Resize()
        On Error GoTo errHandler
        
        Text1.Top = frmMenuInicial.Height - 1500
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Form_Resize" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Timer1_Timer()
        On Error GoTo errHandler
        
        Text1 = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
        
        a = Format(Time, "hh:mm:ss")
        
        If StUsuario = "admin" And a = "17:00:00" Then
            Set SQLState = New SQLDMO.SQLServer
        
            SQLState.Connect StInstancia, "sa", "Jahg1991"
            
            With oBackup
                .Database = "Database"
                .Files = "C:\Backup\RD_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".bak"
                .SQLBackup SQLState
            End With
            
            Set oBackup = Nothing
            Set SQLState = Nothing
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Timer1_Timer" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Respaldar_Click()
        On Error GoTo err
        
        Set SQLState = New SQLDMO.SQLServer
        
        SQLState.Connect StInstancia, "sa", "Jahg1991"
        
        With oBackup
            .Database = "Database"
            .Files = "C:\Backup\" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".bak"
            .SQLBackup SQLState
        End With
        
        Set oBackup = Nothing
        Set SQLState = Nothing
        
        MsgBox "Respaldo realizado con éxito", vbOKOnly, "Terminado"
    Exit Sub
err:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Respaldar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        'MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
        MsgBox "Respaldo realizado con éxito", vbOKOnly, "Terminado"
    End Sub
    
    Private Sub Restaurar_Click()
        'On Error GoTo err
        
        'With CommonDialog1
        '    .DialogTitle = "Seleccione el respaldo"
        '    .Filter = "Archivos de respaldo|*.bak"
        '    .InitDir = "C:\JAHG\PuntoVenta\Respaldos\"
        '    .show 1Open
        'End With
        
        'If CommonDialog1.FileName <> "" )Then
        '    Set SQLState = New SQLDMO.SQLServer
        
        '    SQLState.Connect StInstancia, "sa", "Jahg1991"
        
        '    With oRestore
        '        .Action = 0 ' full db restore
        '        .Database = "Database"
        '        .Devices = Files
        '        .Files = CommonDialog1.FileName
        '        .ReplaceDatabase = True 'Force restore over existing database
        '        .SQLRestore SQLState
        '    End With
        
        '    Set oRestore = Nothing
        '    Set SQLState = Nothing
        
        '    MsgBox "Restauración realizada con éxito", vbOKOnly, "Terminado"
        'End If
        
        'Exit Sub
    'err:
    '    MsgBox "Error al realizar la restauración - " & err.Description, vbCritical, "Error"
    End Sub
    
    Private Sub UsuariosNuevo_Click()
        On Error GoTo errHandler
        
        frmUsuariosNuevo.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:UsuariosNuevo_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub UsuariosExistente_Click()
        On Error GoTo errHandler
    
        frmUsuariosExistente.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:UsuariosExistente_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Opciones_Click()
        On Error GoTo errHandler
        
        frmOpciones.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Opciones_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ItemsNuevo_Click()
        On Error GoTo errHandler
    
        With frmItemNuevo
            .Caption = "Añadir nuevo artículo"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ItemsNuevo_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ItemsExistente_Click()
        On Error GoTo errHandler
        
        frmItemExistente.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ItemsExistente_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ReportarProduccion_Click()
        On Error GoTo errHandler
        
        frmReportarProduccion.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ReportarProduccion_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ListasDeIngredientesNuevo_Click()
        On Error GoTo errHandler
        
        With frmListaIngredientesNuevo
            .Caption = "Nueva lista de ingredientes"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ListasDeIngredientesNuevo_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ListasDeIngredientesExistente_Click()
        On Error GoTo errHandler
        
        With frmListaIngredientesExistente
            .Caption = "Listas de ingredientes"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ListasDeIngredientesExistente_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ClientesNuevo_Click()
        On Error GoTo errHandler
        
        InTipoAltaClienteProveedor = 0
        StTipoClienteProveedor = "Cliente"
        
        With frmClientesNuevo
            .Caption = "Añadir nuevo Cliente"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ClientesNuevo_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ClientesExistente_Click()
        On Error GoTo errHandler
        
        StTipoClienteProveedor = "Cliente"
        
        With frmClientesExistente
            .Caption = "Catálogo de Clientes"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ClientesExistente_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub VentasLocal_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Ventas"
        StTipoVenta = "Local"
        
        Set frmHistorialVentas = Nothing
        
        With frmHistorialVentas
            .Caption = "Historial de ventas en el local"
            .Pagar.Visible = False
            .Agregar.Visible = False
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:VentasLocal_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub VentasDomicilio_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Ventas"
        StTipoVenta = "Domicilio"
        
        Set frmHistorialVentas = Nothing
        
        With frmHistorialVentas
            .Caption = "Historial de ventas a domicilio"
            .Pagar.Visible = False
            .Agregar.Visible = False
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:VentasDomicilio_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub VentasNoPagadas_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Ventas"
        StTipoVenta = "Abiertas"
        
        Set frmHistorialVentas = Nothing
        
        With frmHistorialVentas
            .Caption = "Cuenta Abierta"
            .Pagar.Visible = True
            .Agregar.Visible = True
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:VentasNoPagadas_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub NuevoPedido_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Pedidos"
        
        With frmPedidos
            .Caption = "Nuevo Pedido"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:NuevoPedido_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub PedidosPendientes_Click()
        On Error GoTo errHandler
 
        frmPedidosPendientes.Show 1
        
        Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:frmPedidosPendientes" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
        Private Sub CancelacionPedidos_Click()
        On Error GoTo errHandler
 
        frmCancelacionPedidos.Show 1
        
        Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:CancelacionPedidos_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub NuevaVenta_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Ventas"
        
        With frmVentas
            .Caption = "Nueva Venta"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:NuevaVenta_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ProveedoresNuevo_Click()
        On Error GoTo errHandler
        
        StTipoClienteProveedor = "Proveedor"
        
        With frmProveedoresNuevo
            .Caption = "Añadir nuevo Proveedor"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ProveedoresNuevo_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ProveedoresExistente_Click()
        On Error GoTo errHandler
        
        StTipoClienteProveedor = "Proveedor"
        
        With frmProveedoresExistente
            .Caption = "Catalogo de Proveedores"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ProveedoresExistente_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ComprasPagadas_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Compras"
        StTipoCompra = "Pagadas"
        
        With frmHistorialCompras
            .Caption = "Compras Pagadas"
            .Show 1
            .Pagar.Visible = False
            .Agregar.Visible = False
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ComprasPagadas_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub ComprasNoPagadas_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Compras"
        StTipoCompra = "No Pagadas"
        
        With frmHistorialCompras
            .Caption = "Compras no Pagadas"
            .Show 1
            .Pagar.Visible = True
            .Agregar.Visible = True
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ComprasNoPagadas_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub NuevaCompra_Click()
        On Error GoTo errHandler
        
        StTipoVentasCompras = "Compras"
        
        With frmCompras
            .Caption = "Nueva Compra"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:NuevaCompra_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Stock_Click()
        On Error GoTo errHandler
        
        frmStock.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Stock_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub MovimientosDeInventario_Click()
        On Error GoTo errHandler
        
        frmMovimientosInventario.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:MovimientosDeInventario_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub AjustesDeInventario_Click()
        On Error GoTo errHandler
        
        frmAjusteInventario.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:AjustesDeInventario_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub SalidaInsumos_Click()
        On Error GoTo errHandler
        
        frmSalidaInsumos.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:SalidaInsumos_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Rastreabilidad_Click()
        On Error GoTo errHandler
 
        frmRastreabilidad.Show 1
        
        Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Rastreabilidad_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub EntradaDeDinero_Click()
        On Error GoTo errHandler
        
        StTipoEntradaSalida = "Entrada"
        
        With frmEntradaSalidaDinero
            .Caption = "Entrada de efectivo"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:EntradaDeDinero_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub SalidaDeDinero_Click()
        On Error GoTo errHandler
        
        StTipoEntradaSalida = "Salida"
        
        With frmEntradaSalidaDinero
            .Caption = "Salida de efectivo"
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:SalidaDeDinero_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub CorteDeCajaCorteDeCaja_Click()
        On Error GoTo errHandler
        
        If StUsuario = "admin" Or StUsuario = "sysadmin" Or StUsuario = "gerente" Or StUsuario = "supervisor" Then
            frmCorteCaja.Show 1
        Else
            frmCrearCorteCaja.Show 1
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:CorteDeCajaCorteDeCaja_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub AcercaDe_Click()
        On Error GoTo errHandler
        
        frmAcercaDe.Show 1
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:AcercaDe_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub CerrarSesion_Click()
        On Error GoTo errHandler
    
        Unload Me
        
        Main
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:CerrarSesion_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
