Attribute VB_Name = "Module1"
Global Cn As New ADODB.Connection
Global StConnection As String
Global CmBackup As New ADODB.Command
Global CmRestore As New ADODB.Command
Global StUsuario As String
Global StPermisosArchivo As String
Global StPermisosArticulos As String
Global StPermisosProduccion As String
Global StPermisosVentas As String
Global StPermisosCompras As String
Global StPermisosInventario As String
Global StPermisosCorteCaja As String
Global StTipoClienteProveedor As String
Global StTipoVentasCompras As String
Global StTipoVenta As String
Global StTipoCompra As String
Global StTipoEntradaSalida As String
Global PcNombreEmpresa As String
Global PcRFC As String
Global PcDireccion As String
Global PcTelefono As String
Global InTipoAltaClienteProveedor As Integer
Global IdCliente As Integer

Sub Main()
    
    On Error Resume Next
    
    Dim Rs As New ADODB.Recordset
    
    StConnection = "Provider=SQLOLEDB.1;Password=Jahg1991;Persist Security Info=True;User ID=sa;Initial Catalog=DataBase;Data Source=localhost\SSDB"
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    'With CmBackup
    '    .CommandText = "BACKUP DATABASE [DataBase] TO  DISK = N'" & App.Path & "\Respaldos\Auto_Backup_DataBase" & Replace(Date, "/", "") & ".bck ', WITH NOFORMAT, NOINIT,  NAME = N'SQLShackDemo-Full Database Backup', SKIP, NOREWIND, NOUNLOAD, COMPRESSION, STATS = 10"
    '    .ActiveConnection = Cn
    '    MsgBox .CommandText
    '    '.Execute
    'End With
    
    With Rs
        
        If .State = 1 Then .Close
        .Open "Select * from preferencias;", Cn, adOpenStatic, adLockOptimistic
        .Filter = ""
        .Requery
        
        If .RecordCount <> 0 Then
            .MoveFirst
            PcNombreEmpresa = .Fields(1).Value
            PcRFC = .Fields(2).Value
            PcDireccion = .Fields(3).Value
            PcTelefono = .Fields(4).Value
        End If
        
        .Close
    
    End With
    
    Cn.Close
    
    frmInicioSesion.Show

End Sub

Function Get_ItemId(P_description As String) As Integer
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select Id from MapeoItems where nombre = '" & P_description & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_ItemId = Rs.Fields(0)
    Rs.Close
    
    Exit Function

err:
    
    Get_ItemId = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_ItemCod(P_id As Integer) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select codigo from items where id = " & P_id & ";"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_ItemCod = Rs.Fields(0)
    Rs.Close
    
    Exit Function

err:
    
    Get_ItemCod = ""
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_ItemDesc(P_id As Integer) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select descripcion from items where id = " & P_id & ";"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_ItemDesc = Rs.Fields(0)
    Rs.Close
    
    Exit Function

err:
    
    Get_ItemDesc = ""
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_ItemUDM(P_id As Integer) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select udm from items where id = " & P_id & ";"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_ItemUDM = Rs.Fields(0)
    Rs.Close
    
    Exit Function

err:
    
    Get_ItemUDM = ""
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_ItemTipo(P_id As Integer) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select tipo from items where id = " & P_id & ";"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_ItemTipo = Rs.Fields(0)
    Rs.Close
    
    Exit Function

err:
    
    Get_ItemTipo = ""
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_ItemIva(P_id As Integer) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select iva from items where id = " & P_id & ";"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_ItemIva = Replace(Rs.Fields(0), ",", ".")
    Rs.Close
    
    Exit Function

err:
    
    Get_ItemIva = "0.00"
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_LIExists(P_id As Integer) As Integer
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select count(*) from ListasDeIngredientes where ItemPTId = " & P_id & ";"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_LIExists = Rs.Fields(0)
    Rs.Close
    
    Exit Function

err:
    
    Get_LIExists = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumSubtotal(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(subtotal) from ticket where folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumSubtotal = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumSubtotal = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumIva(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(iva) from ticket where folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumIva = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumIva = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumTotal(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(total) from ticket where folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumTotal = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumTotal = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_Comentario(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select distinct comentarios from HistorialVentasCompras where folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_Comentario = Rs.Fields(0)
    
    Rs.Close

    Exit Function

err:

    Get_Comentario = ""
    If Rs.State = 1 Then Rs.Close

End Function
