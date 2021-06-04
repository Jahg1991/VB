Attribute VB_Name = "Module1"
Global Cn As New ADODB.Connection
Global StConnection As String
Global StUsuario As String
Global StPermisosArchivo As String
Global StPermisosArticulos As String
Global StPermisosVentas As String
Global StPermisosCompras As String
Global StPermisosInventario As String
Global StPermisosCorteCaja As String
Global StTipoItem As String
Global StTipoClienteProveedor As String
Global StTipoVentasCompras As String
Global StTipoVenta As String
Global StTipoCompra As String
Global StTipoEntradaSalida As String
Global StSerieItem As String
Global PcNombreEmpresa As String
Global PcRFC As String
Global PcDireccion As String
Global PcTelefono As String
Global PcImpresoraBarra As String
Global PcImpresoraCocina As String
Global PcImpresoraCompras As String
Global PcImpresoraCorteCaja As String
Global PcNumeroMesas As String
Global InTipoAltaClienteProveedor As Integer
Global IdCliente As Integer
Global EncontroImpresora As Integer

Sub Main()
    
    On Error Resume Next
    
    Dim Rs As New ADODB.Recordset
    
    FileCopy App.Path & "\DataBase.db", App.Path & "\Respaldos\Auto_Backup_DataBase" & Replace(Date, "/", "") & ".bck"
    
    StConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase.db;Persist Security Info=False"
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
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
            PcImpresoraBarra = .Fields(5).Value
            PcImpresoraCocina = .Fields(6).Value
            PcImpresoraCompras = .Fields(7).Value
            PcImpresoraCorteCaja = .Fields(8).Value
            PcNumeroMesas = .Fields(9).Value
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

Function Get_SumSubtotalB(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(subtotal) from ticket where tipoarticulo ='Barra' and folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumSubtotalB = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumSubtotalB = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumIvaB(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(iva) from ticket where tipoarticulo ='Barra' and folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumIvaB = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumIvaB = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumTotalB(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(total) from ticket where tipoarticulo ='Barra' and folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumTotalB = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumTotalB = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumSubtotalC(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(subtotal) from ticket where tipoarticulo <> 'Barra' and folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumSubtotalC = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumSubtotalC = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumIvaC(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(iva) from ticket where tipoarticulo <> 'Barra' and folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumIvaC = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumIvaC = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Get_SumTotalC(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select sum(total) from ticket where tipoarticulo <> 'Barra' and folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_SumTotalC = Replace(Rs.Fields(0), ",", ".")
    Rs.Close

    Exit Function

err:

    Get_SumTotalC = 0
    If Rs.State = 1 Then Rs.Close

End Function

Function Establecer(ByVal NamePrinter As String) As Boolean
  
    On Error GoTo errSub
    EncontroImpresora = 0

    ' Establece la impresora que se utilizará para imprimir
    
    'Variable de referencia
    Dim obj_Impresora As Object
      
    'Creamos la referencia
    Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter NamePrinter
      
    Set obj_Impresora = Nothing
          
        'La función devuelve true y se cambió con éxito
        Establecer = True
    Exit Function
      
      
'Error al cambiar la impresora
errSub:
If err.Number = 0 Then Exit Function
   Establecer = False
   'MsgBox "error: " & err.Number & Chr(13) & "Description: " & err.Description
   EncontroImpresora = 1
   On Error GoTo 0
    
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

Function Get_Mesa(P_folio As String) As String
    
    Dim Rs As New ADODB.Recordset
    Dim St As String
    
    On Error GoTo err
    
    St = "Select distinct mesa from HistorialVentasCompras where folio = '" & P_folio & "';"
    
    With Rs
        If .State = 1 Then .Close
            .Open St, Cn, adOpenStatic, adLockOptimistic
            .Requery
            .MoveFirst
    End With
    
    Get_Mesa = Rs.Fields(0)
    
    Rs.Close

    Exit Function

err:

    Get_Mesa = ""
    If Rs.State = 1 Then Rs.Close

End Function
