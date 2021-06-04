Attribute VB_Name = "Module1"
    Option Explicit
    
    '//CONEXION
    Public Cn                           As New adodb.Connection
    
    '//DATOS DE CONEXION
    Public StInstancia                  As String
    Public StConnection                 As String
    Public StUsuario                    As String
    Public n_File                       As Long
    Public Linea                        As String
    
    '//PERMISOS
    Public StPermisosArchivo            As String
    Public StPermisosArticulos          As String
    Public StPermisosVentas             As String
    Public StPermisosCompras            As String
    Public StPermisosInventario         As String
    Public StPermisosCorteCaja          As String
    Public StPermisosProduccion         As String
    Public StPermisosCaja               As String
    
    '//VARIABLES VENTAS/COMPRAS
    Public StTipoClienteProveedor       As String
    Public StTipoVentasCompras          As String
    Public StTipoVenta                  As String
    Public StTipoCompra                 As String
    Public StTipoEntradaSalida          As String
    
    '//PREFERENCIAS
    Public RsPreferencias               As New adodb.Recordset
    Public PcNombreEmpresa              As String
    Public PcRFC                        As String
    Public PcDireccion                  As String
    Public PcTelefono                   As String
    Public PcValorPuntos                As String
    Public InTipoAltaClienteProveedor   As Long
    Public IdCliente                    As Long
    
    '//FUNCIONES
    Public RsFuncion                    As New adodb.Recordset
    Public StFuncion                    As String
    
    '//RESPALDOS
    Public oBackup                      As New SQLDMO.Backup
    Public SQLState                     As New SQLDMO.SQLServer
    Public oRestore                     As New SQLDMO.Restore
    
    '//IMPRESORA
    Public obj_Impresora                As Object
    
    '//MANEJO DE ERRORES
    Public FileNum                      As Long
    
    ' Constantes para indicar el color de fondo del combobox
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Const COLOR_NO_ENCONTRADO = &HC0C0FF                 ' color cuando no se encontró
    Public Const COLOR_NORMAL = &HE0E0E0                        ' color cuando hay coincidencia
    
    Function Getinstancia() As String
        On Error GoTo errHandler
        
        n_File = FreeFile
        
        Open App.Path & "\instancia" For Input As n_File
        
        Do While Not EOF(n_File)
            Line Input #n_File, Linea
        Loop

        Close n_File
        
        Getinstancia = Linea
    Exit Function
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Getinstancia" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Sub Main()
        On Error GoTo errHandler
        
        StInstancia = Getinstancia()
        Set SQLState = New SQLDMO.SQLServer
        SQLState.Connect StInstancia, "sa", "Jahg1991"
        
        With oBackup
            .Database = "Database"
            .Files = App.Path & "\Respaldos\Automatico_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".bak"
            .SQLBackup SQLState
        End With
        
        Set oBackup = Nothing
        Set SQLState = Nothing
        
        StConnection = "Provider=SQLOLEDB.1;Password=Jahg1991;Persist Security Info=True;User ID=sa;Initial Catalog=DataBase;Data Source=" & StInstancia
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open StConnection
        End With
        
        With RsPreferencias
            If .State = 1 Then .Close
            .Open "Select * from FND_SYSTEM_OPTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            
            If .RecordCount <> 0 Then
                .MoveFirst
                
                If IsNull(.Fields(1).Value) = False Then PcNombreEmpresa = .Fields(1).Value Else PcNombreEmpresa = ""
                If IsNull(.Fields(2).Value) = False Then PcRFC = .Fields(2).Value Else PcRFC = ""
                If IsNull(.Fields(3).Value) = False Then PcDireccion = .Fields(3).Value Else PcDireccion = ""
                If IsNull(.Fields(4).Value) = False Then PcTelefono = .Fields(4).Value Else PcTelefono = ""
                If IsNull(.Fields(5).Value) = False Then PcValorPuntos = Replace(.Fields(5).Value, ",", ".") Else PcValorPuntos = ""
            End If
            
            .Close
        End With
        
        Cn.Close
        
        frmInicioSesion.Show
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Main" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
        
        If err.Number = -2147218303 Then Resume Next
    End Sub
    
    Function Get_ItemId(P_description As String) As Long
        On Error GoTo err
        
        StFuncion = "Select isNull(Id,0) from MTL_SYSTEM_ITEMS_M where nombre = '" & P_description & "';"
        
        Get_ItemId = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_ItemId = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemId" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemCod(P_id As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(codigo,'') from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemCod = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ItemCod = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemCod" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemDesc(P_id As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(descripcion,'') from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemDesc = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ItemDesc = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemDesc" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemUDM(P_id As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(udm,'') from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemUDM = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ItemUDM = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemUDM" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemTipo(P_id As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(tipo,'') from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemTipo = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ItemTipo = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemTipo" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemIva(P_id As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(iva,0) from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemIva = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ItemIva = "0.00"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemIva" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_LIExists(P_id As Long) As Long
        On Error GoTo err
        
        StFuncion = "Select count(*) from BILL_OF_MATERIAL where ItemPTId = " & P_id & ";"
        
        Get_LIExists = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_LIExists = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_LIExists" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_SumSubtotal(P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(sum(subtotal),0) from PO_TRANSACTION_TICKET where folio = '" & P_folio & "';"
        
        Get_SumSubtotal = Replace((CrearFuncionString(StFuncion)), ",", ".")
    Exit Function
err:
        Get_SumSubtotal = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_SumSubtotal" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_SumIva(P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(sum(iva),0) from PO_TRANSACTION_TICKET where folio = '" & P_folio & "';"
        
        Get_SumIva = Replace((CrearFuncionString(StFuncion)), ",", ".")
    Exit Function
err:
        Get_SumIva = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_SumIva" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_SumTotal(P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(sum(total),0) from PO_TRANSACTION_TICKET where folio = '" & P_folio & "';"
        
        Get_SumTotal = Replace((CrearFuncionString(StFuncion)), ",", ".")
    Exit Function
err:
        Get_SumTotal = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_SumTotal" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Establecer(ByVal NamePrinter As String) As Boolean
        On Error GoTo errSub
    
        'Creamos la referencia
        Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter NamePrinter
        Set obj_Impresora = Nothing
              
        'La función devuelve true y se cambió con éxito
        Establecer = True
    Exit Function
errSub:
        If err.Number = 0 Then Exit Function
        
        Establecer = False
       
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Establecer" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_Comentario(P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select distinct isNull(comentarios,'') from PO_LINES_ALL where folio = '" & P_folio & "';"
        
        Get_Comentario = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_Comentario = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_Comentario" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemLote(P_id As Long) As Long
        On Error GoTo err
        
        StFuncion = "Select isNull(lote,0) from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemLote = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_ItemLote = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemLote" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_LoteExiste(P_lote As String, P_item As Long) As Long
        On Error GoTo err
        
        StFuncion = "Select count (*) as existe from MTL_LOT_NUMBERS where item = " & P_item & " and lote = '" & P_lote & "';"
        
        Get_LoteExiste = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_LoteExiste = 1
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_LoteExiste" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_LoteConsumo(P_item As Long) As String
        On Error GoTo err
        
        StFuncion = "Select top 1 isnull(lote,'') as lote from MTL_LOT_ON_HAND_QUANTITIES where item = " & P_item & " order by id;"
        
        Get_LoteConsumo = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_LoteConsumo = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_LoteConsumo" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_LoteConsumoCantidad(P_item As Long) As String
        On Error GoTo err
        
        StFuncion = "Select top 1 isNull(cantidad,0) as cantidad from MTL_LOT_ON_HAND_QUANTITIES where item = " & P_item & " order by id;"
        
        Get_LoteConsumoCantidad = Replace(Format(CrearFuncionString(StFuncion), "0.00"), ",", ".")
    Exit Function
err:
        Get_LoteConsumoCantidad = "0.00"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_LoteConsumoCantidad" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_Credito(P_Cliente As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(credito,0) from HZ_PARTY where id = " & P_Cliente & " ;"
        
        Get_Credito = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_Credito = "0"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_Credito" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_CreditoDias(P_Cliente As Long) As Long
        On Error GoTo err
        
        StFuncion = "Select isNull(credito_dias,0) from HZ_PARTY where id = " & P_Cliente & " ;"
        
        Get_CreditoDias = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_CreditoDias = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_CreditoDias" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_CreditoUsado(P_Cliente As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(credito,0) from HZ_PARTY_CREDIT where idclienteproveedor = " & P_Cliente & " ;"
        
        Get_CreditoUsado = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_CreditoUsado = "0"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_CreditoUsado" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_CreditoDiasUsado(P_Cliente As Long) As Long
        On Error GoTo err
        
        StFuncion = "Select isNull(dias_credito,0) from HZ_PARTY_CREDIT_DAYS where idclienteproveedor = " & P_Cliente & " ;"
        
        Get_CreditoDiasUsado = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_CreditoDiasUsado = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_CreditoDiasUsado" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ClienteMayorista(P_Cliente As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(mayorista,'No') from HZ_PARTY where id = " & P_Cliente & " ;"
        
        Get_ClienteMayorista = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ClienteMayorista = "No"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ClienteMayorista" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ClienteFolio(P_folio As String) As Long
        On Error GoTo err
        
        StFuncion = "Select Distinct isNull(IdClienteProveedor,0) from PO_LINES_ALL where Folio = '" & P_folio & "';"
        
        Get_ClienteFolio = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_ClienteFolio = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ClienteFolio" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ClienteListaP(P_Cliente As Long) As Long
        On Error GoTo err
        
        StFuncion = "Select isNull([Lista de Precios],0) FROM HZ_PARTY WHERE Id =" & P_Cliente & ";"
        
        Get_ClienteListaP = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_ClienteListaP = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ClienteListaP" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ClientePuntos(P_Cliente As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(puntos,0) From RA_POINT_TRANSACTIONS_V WHERE Cliente =" & P_Cliente & ";"
        
        Get_ClientePuntos = Replace(Format(CrearFuncionString(StFuncion), "0.00"), ",", ".")
    Exit Function
err:
        Get_ClientePuntos = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ClientePuntos" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_DevItemExiste(P_item As Long, P_folio As String) As Long
        On Error GoTo err
        
        StFuncion = "Select count(*) as numero From PO_LINES_ALL WHERE IdArticulo =" & P_item & " and Folio = '" & P_folio & "';"
        
        Get_DevItemExiste = CrearFuncionLong(StFuncion)
    Exit Function
err:
        Get_DevItemExiste = 0
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_DevItemExiste" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_ItemCategoria(P_id As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(categoria,'Gasto') from MTL_SYSTEM_ITEMS where id = " & P_id & ";"
        
        Get_ItemCategoria = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_ItemCategoria = "Gasto"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_ItemCategoria" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_LoteConsumoDev(P_item As Long, P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select top 1 isNull(lote,'') as lote from PO_LOT_ON_HAND_QUANTITIES where item = " & P_item & " and folio = '" & P_folio & "' order by id desc;"
        
        Get_LoteConsumoDev = CrearFuncionString(StFuncion)
    Exit Function
err:
        Get_LoteConsumoDev = ""
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_LoteConsumoDev" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_LoteConsumoCantidadDev(P_item As Long, P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select top 1 isNull(cantidad,0) as cantidad from PO_LOT_ON_HAND_QUANTITIES where item = " & P_item & " and folio = '" & P_folio & "' order by id desc;"
        
        Get_LoteConsumoCantidadDev = Replace(Format(CrearFuncionString(StFuncion), "0.00"), ",", ".")
    Exit Function
err:
        Get_LoteConsumoCantidadDev = "0.00"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_LoteConsumoCantidadDev" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_CantidadDev(P_item As Long, P_folio As String) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(SUM(Cantidad),0) AS cantidad from PO_LINES_ALL where IdArticulo = " & P_item & " and folio = '" & P_folio & "' and Cancelado = 'No';"
        
        Get_CantidadDev = Replace(Format(CrearFuncionString(StFuncion), "0.00"), ",", ".")
    Exit Function
err:
        Get_CantidadDev = "0.00"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_CantidadDev" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function Get_CantidadItem(P_item As Long) As String
        On Error GoTo err
        
        StFuncion = "Select isNull(SUM(Disponible),0) AS cantidad from MTL_ON_HAND_QUANTITIES where itemid = " & P_item & ";"
        
        Get_CantidadItem = Replace(Format(CrearFuncionString(StFuncion), "0.00"), ",", ".")
    Exit Function
err:
        Get_CantidadItem = "0.00"
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: Get_CantidadItem" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function CrearFuncionLong(P_Consulta As String) As Long
        On Error GoTo err
        
        With RsFuncion
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open P_Consulta, Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            
            If .RecordCount <> 0 Then
                .MoveFirst
            
                CrearFuncionLong = .Fields(0).Value
            Else
                CrearFuncionLong = 0
            End If
        
            .Close
        End With
    Exit Function
err:
        CrearFuncionLong = 0
        
        If RsFuncion.State = 1 Then RsFuncion.Close
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: CrearFuncionLong" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
    
    Function CrearFuncionString(P_Consulta As String) As String
        On Error GoTo err
        
        With RsFuncion
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open P_Consulta, Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            
             If .RecordCount <> 0 Then
                .MoveFirst
            
                CrearFuncionString = .Fields(0).Value
            Else
                CrearFuncionString = ""
            End If
        
            .Close
        End With
    Exit Function
err:
        CrearFuncionString = ""
        
        If RsFuncion.State = 1 Then RsFuncion.Close
        
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: CrearFuncionString" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Function
