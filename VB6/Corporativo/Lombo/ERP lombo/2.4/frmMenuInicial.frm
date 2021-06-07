VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenuInicial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Inicial"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   17310
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
   Icon            =   "frmMenuInicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmMenuInicial.frx":000C
   ScaleHeight     =   4185
   ScaleWidth      =   17310
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
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
      Height          =   465
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   3240
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
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
      Picture         =   "frmMenuInicial.frx":E383
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10335
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Opciones 
         Caption         =   "Opciones"
      End
      Begin VB.Menu Respaldar 
         Caption         =   "Respaldar"
      End
      Begin VB.Menu Restaurar 
         Caption         =   "Restaurar"
      End
      Begin VB.Menu Usuarios 
         Caption         =   "Usuarios"
         Begin VB.Menu UsuariosNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu UsuariosExistente 
            Caption         =   "Existente"
         End
      End
   End
   Begin VB.Menu Catalogos 
      Caption         =   "Catalogos"
      Begin VB.Menu Articulos 
         Caption         =   "Artículos"
         Begin VB.Menu ItemsNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu ItemsExistente 
            Caption         =   "Existente"
         End
         Begin VB.Menu CategoriasArticulosNuevo 
            Caption         =   "Categorias"
         End
      End
      Begin VB.Menu Clientes 
         Caption         =   "Clientes"
         Begin VB.Menu ClientesNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu ClientesExistente 
            Caption         =   "Existente"
         End
      End
      Begin VB.Menu Proveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu ProveedoresNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu ProveedoresExistente 
            Caption         =   "Existente"
         End
         Begin VB.Menu CategoriaProveedoresNuevo 
            Caption         =   "Categorias"
         End
      End
   End
   Begin VB.Menu ListasDeIngredientes 
      Caption         =   "Lista de Materiales"
      Begin VB.Menu ListasDeIngredientesNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu ListasDeIngredientesExistente 
         Caption         =   "Existente"
      End
   End
   Begin VB.Menu Produccion 
      Caption         =   "Produccion"
      Begin VB.Menu ReportarProduccion 
         Caption         =   "Reportar Produccion"
      End
   End
   Begin VB.Menu Ventas 
      Caption         =   "Ventas"
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
      End
      Begin VB.Menu NuevaVenta 
         Caption         =   "Nueva Venta"
      End
   End
   Begin VB.Menu Pedidos 
      Caption         =   "Pedidos"
      Begin VB.Menu NuevoPedido 
         Caption         =   "Nuevo Pedido"
      End
      Begin VB.Menu PedidosPendientes 
         Caption         =   "Pedidos Pendientes"
      End
      Begin VB.Menu CancelacionPedidos 
         Caption         =   "Cancelacion de Pedidos"
      End
   End
   Begin VB.Menu Compras 
      Caption         =   "Compras"
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
      End
   End
   Begin VB.Menu Ajustes 
      Caption         =   "Ajustes de Inventario"
      Begin VB.Menu AjustesDeInventario 
         Caption         =   "Realizar Ajuste"
      End
   End
   Begin VB.Menu Inventario 
      Caption         =   "Inventario"
      Begin VB.Menu SalidaInsumos 
         Caption         =   "Salida de Insumos"
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
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu rCatalogos 
         Caption         =   "Catalogos"
      End
      Begin VB.Menu rListasDeIngredientes 
         Caption         =   "Lista de Materiales"
      End
      Begin VB.Menu rProduccion 
         Caption         =   "Producción"
      End
      Begin VB.Menu rVentas 
         Caption         =   "Ventas"
      End
      Begin VB.Menu rPedidos 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu rCompras 
         Caption         =   "Compras"
      End
      Begin VB.Menu rInventarios 
         Caption         =   "Inventarios"
         Begin VB.Menu Stock 
            Caption         =   "Cantidad en Mano"
         End
         Begin VB.Menu MovimientosDeInventario 
            Caption         =   "Movimientos de Inventario"
         End
         Begin VB.Menu Rastreabilidad 
            Caption         =   "Rastreabilidad"
         End
         Begin VB.Menu rCorteDeCaja 
            Caption         =   "Corte de Caja"
         End
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
'***********************************************************************************
'Nombre:        frmMenuInicial
'Proposito:     Menu con las operaciones disponibles para el usuario
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Se agregó manejo para error
'                                               -2147217900 en la restauracion.
'                                               El respaldo no es valido
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//OTROS
Dim i As Long
Dim a As String

Private Sub Form_Load()
    On Error GoTo errHandler
    With Combo1
        For i = 1 To 10
            .AddItem "Caja " & i
        Next i
        .Text = StCajaPredeterminada
    End With

    With Image1
        .Picture = LoadPicture(App.Path & "\Images\Inicio.jpg")
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
    With Text1
        .Top = frmMenuInicial.Height - 1500
    End With
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
    With Text1
        .Text = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
    End With
    a = Format(Time, "hh:mm:ss")
    If StUsuario = "admin" And a = "17:00:00" Then
        Set SQLState = New SQLDMO.SQLServer

        With SQLState
            .Connect StInstancia, "sa", "Jahg1991"
        End With

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

'**********************************************
'ARCHIVO
'**********************************************
Private Sub Opciones_Click()
    On Error GoTo errHandler
    With frmOpciones
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Opciones_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Respaldar_Click()
    On Error GoTo err
    Set SQLState = New SQLDMO.SQLServer

    With SQLState
        .Connect StInstancia, "sa", "Jahg1991"
    End With

    With oBackup
        .Database = "Database"
        .Files = StRespaldo & "\" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".bak"
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
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Restaurar_Click()
    On Error GoTo err
    With CommonDialog1
        .FileName = ""
        .DialogTitle = "Seleccione el respaldo"
        .Filter = "Archivos de respaldo|*.bak"
        .InitDir = StRespaldo & "\"
        .ShowOpen
        If .FileName <> "" Then
            With Cn
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                If .State = 0 Then .Open (StConnection)
                .Execute "USE [master]"
                .Execute "RESTORE DATABASE [Database] FROM DISK='" & CommonDialog1.FileName & "'"
                .Close
            End With

            Set Cn = Nothing
            MsgBox "Restauración realizada con éxito", vbOKOnly, "Terminado"
        End If
    End With
    Exit Sub
err:
    If err.Number = -2147217900 Then
        Set Cn = Nothing
        MsgBox "El archivo de respaldo no es válido", vbOKOnly, "Error"
        err.Clear
        Exit Sub
    End If
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Restaurar_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub UsuariosNuevo_Click()
    On Error GoTo errHandler
    With frmUsuariosNuevo
        .Show 1
    End With
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
    With frmUsuariosExistente
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:UsuariosExistente_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'CATALOGOS
'**********************************************
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
    With frmItemExistente
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ItemsExistente_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub CategoriasArticulosNuevo_Click()
    On Error GoTo errHandler
    With frmCategoriaArticuloNuevo
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:CategoriasArticulosNuevo_Click" & vbTab & err.Number & vbTab & err.Description
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

Private Sub CategoriaProveedoresNuevo_Click()
    On Error GoTo errHandler
    With frmCategoriaProveedorNuevo
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:CategoriaProveedoresNuevo_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'LISTA DE MATERIALES
'**********************************************
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

'**********************************************
'PRODUCCION
'**********************************************
Private Sub ReportarProduccion_Click()
    On Error GoTo errHandler
    With frmReportarProduccion
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:ReportarProduccion_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'VENTAS
'**********************************************
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
        With .Pagar
            .Visible = False
        End With

        With .Agregar
            .Visible = False
        End With
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
        With .Pagar
            .Visible = True
        End With

        With .Agregar
            .Visible = True
        End With
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

'**********************************************
'PEDIDOS
'**********************************************
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
    With frmPedidosPendientes
        .Show 1
    End With
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
    With frmCancelacionPedidos
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:CancelacionPedidos_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'COMPRAS
'**********************************************
Private Sub ComprasPagadas_Click()
    On Error GoTo errHandler
    StTipoVentasCompras = "Compras"
    StTipoCompra = "Pagadas"
    With frmHistorialCompras
        .Caption = "Compras Pagadas"
        .Show 1
        With .Pagar
            .Visible = False
        End With

        With .Agregar
            .Visible = False
        End With
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
        With .Pagar
            .Visible = True
        End With

        With .Agregar
            .Visible = True
        End With
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

'**********************************************
'AJUSTES
'**********************************************
Private Sub AjustesDeInventario_Click()
    On Error GoTo errHandler
    With frmAjusteInventario
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:AjustesDeInventario_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'INSUMOS
'**********************************************
Private Sub SalidaInsumos_Click()
    On Error GoTo errHandler
    With frmSalidaInsumos
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:SalidaInsumos_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'CAJAS
'**********************************************
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
        With frmCorteCaja
            .Show 1
        End With
    Else
        With frmCrearCorteCaja
            .Show 1
        End With
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

'**********************************************
'REPORTES
'**********************************************
Private Sub Stock_Click()
    On Error GoTo errHandler
    With frmStock
        .Show 1
    End With
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
    With frmMovimientosInventario
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:MovimientosDeInventario_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Rastreabilidad_Click()
    On Error GoTo errHandler
    With frmRastreabilidad
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:Rastreabilidad_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'ACERCA DE
'**********************************************
Private Sub AcercaDe_Click()
    On Error GoTo errHandler
    With frmAcercaDe
        .Show 1
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmMenuInicial:AcercaDe_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

'**********************************************
'CERRAR SESION
'**********************************************
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

'**********************************************
'SALIR
'**********************************************
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
