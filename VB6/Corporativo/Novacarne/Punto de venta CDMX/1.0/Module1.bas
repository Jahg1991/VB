Attribute VB_Name = "Module1"
Global Cn As New ADODB.Connection

Global CnString As String

Global RsIdVenta As New ADODB.Recordset
Global RsClientes As New ADODB.Recordset
Global RsItems As New ADODB.Recordset
Global RsPagos As New ADODB.Recordset
Global RsVentas As New ADODB.Recordset
Global RsPagosV As New ADODB.Recordset
Global RsSaldosV As New ADODB.Recordset
Global RsVentasV As New ADODB.Recordset
Global RsVentasTxt As New ADODB.Recordset
Global RsCabeceraVentas As New ADODB.Recordset
Global RsIdPagos As New ADODB.Recordset
Global RsPreferencias As New ADODB.Recordset

Global StIdVenta As String
Global STClientes As String
Global StItems As String
Global StPagos As String
Global StVentas As String
Global StPagosV As String
Global StSaldosV As String
Global StVentasV As String
Global StExportVentas As String
Global StCabeceraVentas As String
Global StIdPagos As String
Global StPreferencias As String

Global IdTransaccion As Integer
Global IdPagos As Integer
Global IdItem As String
Global IdCliente As String

Global TipoCatalogo As Integer
    '   Valor   Significado
    '   0       Articulos
    '   1       Clientes
    '   2       Pagos
    '   3       Ventas

Global TipoConsulta As Integer
    '   Valor   Significado
    '   0       Consulta
    '   1       Seleccion

Global TipoTransaccion As String
    '   Valor   Significado
    '   0       Pago
    '   1       Venta
    
Global PrEmpresa As String
Global PrRfc As String
Global PrCiudad As String
Global PrEstado As String
Global PrCodigoPostal As String
Global PrTelefono As String
Global PrCorreo As String
Global PrCliente As Integer
Global PrArticulo As Integer
Global PrPrecio As Integer
Global PrTicket As Integer
Global PrPreferencias As Integer
Global PrImpresora As String
Global PrERP As Integer

Sub OpenBd()

CnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\BD.mdb;" & _
           "Persist Security Info=False"
    With Cn
        .CursorLocation = adUseClient
        .Open CnString
    End With

End Sub

Sub CargarRs()

    StIdVenta = "Select * from IdTransaccion"
    STClientes = "Select * from Clientes"
    StItems = "Select * from Items"
    StPagos = "Select * from Pagos"
    StVentas = "Select * from Ventas"
    StPagosV = "Select * from Pagos_v"
    StSaldosV = "Select * from Saldos_v"
    StVentasV = "Select * from Ventas_v"
    StExportVentas = "Select * from export_ventas"
    StCabeceraVentas = "Select * from CabeceraVentas_v"
    StIdPagos = "Select * from IdPagos"
    StPreferencias = "Select * from Preferencias"
    With RsIdVenta
        If .State = 1 Then .Close
            .Open StIdVenta, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsClientes
        If .State = 1 Then .Close
            .Open STClientes, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsItems
        If .State = 1 Then .Close
            .Open StItems, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsPagos
        If .State = 1 Then .Close
            .Open StPagos, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsVentas
        If .State = 1 Then .Close
            .Open StVentas, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsPagosV
        If .State = 1 Then .Close
            .Open StPagosV, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsSaldosV
        If .State = 1 Then .Close
            .Open StSaldosV, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsVentasV
        If .State = 1 Then .Close
            .Open StVentasV, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsVentasTxt
        If .State = 1 Then .Close
            .Open StExportVentas, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsCabeceraVentas
        If .State = 1 Then .Close
            .Open StCabeceraVentas, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsIdPagos
        If .State = 1 Then .Close
            .Open StIdPagos, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsPreferencias
        If .State = 1 Then .Close
            .Open StPreferencias, Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
End Sub

Sub InitForm()

    CargarRs
    
End Sub

Sub main()

    OpenBd
    InitForm
    CargarPreferencias
    Form1.Show
    
End Sub

Sub CargarPreferencias()

    With RsPreferencias
        .Requery
        .MoveFirst
        If IsNull(.Fields("Empresa")) = False Then
            PrEmpresa = .Fields("Empresa")
        Else
            PrEmpresa = ""
        End If
        If IsNull(.Fields("RFC")) = False Then
            PrRfc = .Fields("RFC")
        Else
            PrRfc = ""
        End If
        If IsNull(.Fields("Ciudad")) = False Then
            PrCiudad = .Fields("Ciudad")
        Else
            PrCiudad = ""
        End If
        If IsNull(.Fields("Estado")) = False Then
            PrEstado = .Fields("Estado")
        Else
            PrEstado = ""
        End If
        If IsNull(.Fields("Codigo Postal")) = False Then
            PrCodigoPostal = .Fields("Codigo Postal")
        Else
            PrCodigoPostal = ""
        End If
        If IsNull(.Fields("Telefono")) = False Then
            PrTelefono = .Fields("Telefono")
        Else
            PrTelefono = ""
        End If
        If IsNull(.Fields("Correo")) = False Then
            PrCorreo = .Fields("Correo")
        Else
            PrCorreo = ""
        End If
        PrCliente = .Fields("Añadir Clientes")
        PrArticulo = .Fields("Añadir Articulos")
        PrPrecio = .Fields("Modificar Precio Venta")
        PrTicket = .Fields("Imprimir Ticket")
        PrPreferencias = .Fields("Ver Preferencias")
        If IsNull(.Fields("Impresora de Tickets")) = False Then
            PrImpresora = .Fields("Impresora de Tickets")
        Else
            PrImpresora = ""
        End If
        PrERP = .Fields("ERP")
    End With

End Sub
