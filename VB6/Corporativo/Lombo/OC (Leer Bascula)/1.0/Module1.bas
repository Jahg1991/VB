Attribute VB_Name = "Module1"
Global CN As New ADODB.Connection
Global RS1 As New ADODB.Recordset
Global RS2 As New ADODB.Recordset
Global RS3 As New ADODB.Recordset
Global RS4 As New ADODB.Recordset
Global IdProveedor As Integer
Global IdArticulo As Integer
Global Precio

Function LeerPuertoBascula() As String

    On Error Resume Next

    Dim cBuffer As String
    
    With Form1.MSComm1
        If .PortOpen = True Then .PortOpen = False
            
            .CommPort = 3 'Numero de puerto que deseas capturar, puede ser cualquier otro numero
            .Settings = "9600,N,8,1"
            .InputLen = 0 'Leer todos los datos
            .InputMode = comInputModeText 'Los datos se dan en modo texto
            .Handshaking = 0
            .PortOpen = True
            
            'limpiamos la variable que almacenara el peso que envie la bascula
            cBuffer = ""
            'En las basculas TORREY debes enviar el caracter 'P' para que te devuelva el peso, en este caso lo envio con Chr$(80)
            .Output = Chr$(80)
            
            'En este ciclo esta el truco para que tome la lectura de la bascula
            Do
            DoEvents
            cBuffer = cBuffer & .Input
            Loop Until InStr(cBuffer, "kg")
            'cerramos el puerto
            .PortOpen = False
    End With
    
    LeerPuertoBascula = cBuffer
    Form1.Text1(0).Text = LeerPuertoBascula
End Function

Sub Main()
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BD.mdb;Persist Security Info=False"
    End With
    Form1.Show
End Sub

Sub CProveedores()
    On Error Resume Next
    Form2.Show
    With Form2
        .Text2.Visible = True
        .Text1.Visible = False
        With RS2
            .Requery
        End With
        With .DataGrid1
            Set .DataSource = RS2
            .Columns(1).Width = 11000
        End With
    End With
End Sub

Sub CArticulos()
    On Error Resume Next
    Form2.Show
    With Form2
        .Text2.Visible = False
        .Text1.Visible = True
        With RS1
            .Requery
        End With
        With .DataGrid1
            Set .DataSource = RS1
            .Columns(1).Width = 8300
        End With
    End With
End Sub

Sub ACatalogo()
    On Error Resume Next
    If Form2.Text2.Visible = True Then 'Proveedores
        RS2.Update
        With Form2.DataGrid1
            .AllowAddNew = False
            .AllowDelete = False
            .Refresh
        End With
        RS2.MoveFirst
    End If
    If Form2.Text1.Visible = True Then 'Artìculos
        RS1.Update
        With Form2.DataGrid1
            .AllowAddNew = False
            .AllowDelete = False
            .Refresh
        End With
        RS1.MoveFirst
    End If
End Sub

Sub ICatalogo()
    On Error Resume Next
    If Form2.Text2.Visible = True Then 'Proveedores
        With Form2.DataGrid1
            .AllowAddNew = True
            .Refresh
        End With
        RS2.MoveLast
    End If
    If Form2.Text1.Visible = True Then 'Artìculos
        With Form2.DataGrid1
            .AllowAddNew = True
            .Refresh
        End With
        RS1.MoveLast
    End If
End Sub

Sub SCatalogo()
    On Error Resume Next
    If Form2.Text2.Visible = True Then 'Proveedores
        If Form2.DataGrid1.AllowAddNew = True Then
            Form2.DataGrid1.AllowAddNew = False
        End If
        RS2.Delete
        RS2.MoveFirst
    End If
    If Form2.Text1.Visible = True Then 'Artìculos
        If Form2.DataGrid1.AllowAddNew = True Then
            Form2.DataGrid1.AllowAddNew = False
        End If
        RS1.Delete
        RS1.MoveFirst
    End If
End Sub

Sub BCatalogo()
    On Error Resume Next
    If Form2.Text1.Visible = True Then
        On Error Resume Next
        With RS1
            .Requery
            If Form2.Text1 <> "" Then
                .Filter = "NOMBRE LIKE '*" & Form2.Text1 & "*'"
                Form2.DataGrid1.Columns(1).Width = 8300
            Else
                .Filter = ""
                Set Form2.DataGrid1.DataSource = RS1
                Form2.DataGrid1.Columns(1).Width = 8300
                .MoveFirst
            End If
        End With
    End If
    If Form2.Text2.Visible = True Then
        On Error Resume Next
        With RS2
            .Requery
            If Form2.Text2 <> "" Then
                .Filter = "NOMBRE LIKE '*" & Form2.Text2 & "*'"
                Form2.DataGrid1.Columns(1).Width = 11000
            Else
                .Filter = ""
                Set Form2.DataGrid1.DataSource = RS2
                Form2.DataGrid1.Columns(1).Width = 11000
                .MoveFirst
            End If
        End With
    End If
End Sub

Sub CExportarExcel()

    On Error GoTo err

    Form2.CommonDialog1.DialogTitle = "Guardar como"
    Form2.CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    Form2.CommonDialog1.ShowSave
    If Form2.CommonDialog1.FileName = "" Then
        MsgBox "Selecciona donde quieres guardar el archivo", vbOKOnly, "Atención"
    End If

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim exportFileName As String
    Dim i As Long
    Dim j As Long
    Dim numFilas As Long
    i = 1
    j = 1
    numFilas = 0
    exportFileName = Form2.CommonDialog1.FileName
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    xlApp.DisplayAlerts = False
    If Form2.Text1.Visible = True Then
        While (numFilas < RS1.Fields.Count)
            xlSheet.Cells(i, j) = RS1.Fields(numFilas).Name
            j = j + 1
            numFilas = numFilas + 1
        Wend
            i = i + 1
        While (Not RS1.EOF)
            j = 0
            While (j < numFilas)
                xlSheet.Cells(i, j + 1).Value = RS1(j)
                j = j + 1
            Wend
            i = i + 1
            RS1.MoveNext
        Wend
    End If
    If Form2.Text2.Visible = True Then
        While (numFilas < RS2.Fields.Count)
            xlSheet.Cells(i, j) = RS2.Fields(numFilas).Name
            j = j + 1
            numFilas = numFilas + 1
        Wend
            i = i + 1
        While (Not RS2.EOF)
            j = 0
            While (j < numFilas)
                xlSheet.Cells(i, j + 1).Value = RS2(j)
                j = j + 1
            Wend
            i = i + 1
            RS2.MoveNext
        Wend
    End If
    xlSheet.SaveAs exportFileName
    xlBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    MsgBox "Archivo creado correctamente"
    Form2.CommonDialog1.FileName = ""
    Exit Sub
err:
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        MsgBox "El archivo no ha podido ser creado" & vbNewLine & err.Description
End Sub

Sub MBProveedor()
    On Error Resume Next
    With Form2
        .Show
        .Text2.Visible = True
        .Text1.Visible = False
        .Command1.Visible = False
        .Command2.Visible = False
        .Command3.Visible = False
        .Command4.Visible = True
        .Exportar.Visible = False
        .DataGrid1.AllowUpdate = False
        Set .DataGrid1.DataSource = RS2
        Form2.DataGrid1.Columns(1).Width = 11000
        RS2.MoveFirst
    End With
End Sub

Sub MBArticulo()
    On Error Resume Next
    With Form2
        .Show
        .Text1.Visible = True
        .Text2.Visible = False
        .Command1.Visible = False
        .Command2.Visible = False
        .Command3.Visible = False
        .Command4.Visible = True
        .Exportar.Visible = False
        .DataGrid1.AllowUpdate = False
        Set .DataGrid1.DataSource = RS1
        Form2.DataGrid1.Columns(1).Width = 8300
        RS1.MoveFirst
    End With
End Sub

Sub BProveedor()
    On Error Resume Next
    If Form2.Text2.Visible = True Then
        With RS2
            .Requery
            If Form2.Text2 <> "" Then
                .Filter = "NOMBRE LIKE '*" & Form2.Text2 & "*'"
                Form2.DataGrid1.Columns(1).Width = 11000
            Else
                .Filter = ""
                Set Form2.DataGrid1.DataSource = RS2
                Form2.DataGrid1.Columns(1).Width = 11000
                .MoveFirst
            End If
        End With
    End If
    If Form2.Text1.Visible = True Then
        With RS1
            .Requery
            If Form2.Text1 <> "" Then
                .Filter = "NOMBRE LIKE '*" & Form2.Text1 & "*'"
                Form2.DataGrid1.Columns(1).Width = 8300
            Else
                .Filter = ""
                Set Form2.DataGrid1.DataSource = RS1
                Form2.DataGrid1.Columns(1).Width = 8300
                .MoveFirst
            End If
        End With
    End If
End Sub

Sub SBProveedor()
    On Error Resume Next
    If Form2.Text2.Visible = True Then
        With RS2
            IdProveedor = .Fields("ID")
            Form1.Text2(0).Text = .Fields("NOMBRE")
        End With
    End If
    If Form2.Text1.Visible = True Then
        With RS1
            IdArticulo = .Fields("ID")
            Precio = .Fields("PRECIO")
            Form1.Text2(1).Text = .Fields("NOMBRE")
            Form1.Text2(2).Text = .Fields("PRECIO")
        End With
    End If
    Unload Form2
End Sub

Sub CargarRS()
    With RS1
        If .State = 1 Then .Close
            .Open "Select * from articulos", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RS2
        If .State = 1 Then .Close
            .Open "Select * from proveedores", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RS3
        If .State = 1 Then .Close
            .Open "Select * from compras", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RS4
        If .State = 1 Then .Close
            .Open "Select * from r_compras", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
End Sub

Sub CTotal()
    On Error Resume Next
    With Form1
        If .Text1(0) <> "" Then
            .Text2(3) = Round(Replace(.Text1(0), "kg", "") * .Text2(2), 2)
        End If
    End With
End Sub

Sub ICompra()
    On Error Resume Next
    With RS3
        .AddNew
            .Fields("ID_PROVEEDOR") = IdProveedor
            .Fields("ID_ARTICULO") = IdArticulo
            .Fields("PRECIO") = Form1.Text2(2)
            .Fields("PESO") = Form1.Text1(0)
            .Fields("TOTAL") = Form1.Text2(3)
            .Fields("FECHA") = Date
            .Fields("TICKET") = Form1.Text2(4)
            .Update
    End With
    
    IdProveedor = Null
    IdArticulo = Null
    Form1.Text2(0) = ""
    Form1.Text2(1) = ""
    Form1.Text2(2) = ""
    Form1.Text2(3) = ""
    Form1.Text2(4) = ""
    Form1.Text1(0) = ""
    MsgBox "Compra guardada", vbOKOnly, "Terminado"
End Sub

Sub F3Load()
    On Error Resume Next
    With RS4
        Set Form3.DataGrid1.DataSource = RS4
        .Requery
        .MoveFirst
    End With
    With Form3
        With .DataGrid1
            .Columns(0).Width = 1500
            .Columns(1).Width = 2500
            .Columns(2).Width = 2500
            .Columns(3).Width = 2000
            .Columns(4).Width = 1500
            .Columns(5).Width = 2500
            .Columns(6).Width = 1500
        End With
    End With
End Sub

Sub CExportarExcelCompras()

    On Error GoTo err

    Form3.CommonDialog1.DialogTitle = "Guardar como"
    Form3.CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    Form3.CommonDialog1.ShowSave
    If Form3.CommonDialog1.FileName = "" Then
        MsgBox "Selecciona donde quieres guardar el archivo", vbOKOnly, "Atención"
    End If

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim exportFileName As String
    Dim i As Long
    Dim j As Long
    Dim numFilas As Long
    i = 1
    j = 1
    numFilas = 0
    exportFileName = Form3.CommonDialog1.FileName
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    xlApp.DisplayAlerts = False
    While (numFilas < RS4.Fields.Count)
        xlSheet.Cells(i, j) = RS4.Fields(numFilas).Name
        j = j + 1
        numFilas = numFilas + 1
    Wend
        i = i + 1
    While (Not RS4.EOF)
        j = 0
        While (j < numFilas)
            xlSheet.Cells(i, j + 1).Value = RS4(j)
            j = j + 1
        Wend
        i = i + 1
        RS4.MoveNext
    Wend
    xlSheet.SaveAs exportFileName
    xlBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    MsgBox "Archivo creado correctamente"
    Form3.CommonDialog1.FileName = ""
    Exit Sub
err:
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        MsgBox "El archivo no ha podido ser creado" & vbNewLine & err.Description
End Sub

Sub LCompra()
    On Error Resume Next
    With Form3
        Set .DataGrid1.DataSource = RS4
        .Text1 = ""
        .Text2 = ""
        .DTPicker1 = Date
        .DTPicker2 = Date
    End With
    With RS4
        .Requery
        .Filter = ""
        .MoveFirst
    End With
    With Form3
        With .DataGrid1
            .Columns(0).Width = 1500
            .Columns(1).Width = 2500
            .Columns(2).Width = 2500
            .Columns(3).Width = 2000
            .Columns(4).Width = 1500
            .Columns(5).Width = 2500
            .Columns(6).Width = 1500
        End With
    End With
    With Form3
        Set .DataGrid1.DataSource = RS4
        .Text1 = ""
        .Text2 = ""
        .DTPicker1 = Date
        .DTPicker2 = Date
    End With
    With RS4
        .Requery
        .Filter = ""
        .MoveFirst
    End With
    With Form3
        With .DataGrid1
            .Columns(0).Width = 1500
            .Columns(1).Width = 2500
            .Columns(2).Width = 2500
            .Columns(3).Width = 2000
            .Columns(4).Width = 1500
            .Columns(5).Width = 2500
            .Columns(6).Width = 1500
        End With
    End With
End Sub

Sub BVentas()
    
    On Error Resume Next
        If Form3.Text1 <> "" And Form3.Text2 <> "" Then
            With RS4
                .Requery
                .Filter = "PROVEEDOR LIKE '*" & Form3.Text1 & "*'" & " AND ARTICULO LIKE '*" & Form3.Text2 & "*'" & " AND FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
            With Form3
                With .DataGrid1
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 2500
                    .Columns(3).Width = 2000
                    .Columns(4).Width = 1500
                    .Columns(5).Width = 2500
                    .Columns(6).Width = 1500
                End With
            End With
        End If
        If Form3.Text1 = "" And Form3.Text2 <> "" Then
            With RS4
                .Requery
                .Filter = "ARTICULO LIKE '*" & Form3.Text2 & "*'" & " AND FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
            With Form3
                With .DataGrid1
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 2500
                    .Columns(3).Width = 2000
                    .Columns(4).Width = 1500
                    .Columns(5).Width = 2500
                    .Columns(6).Width = 1500
                End With
            End With
        End If
        If Form3.Text1 <> "" And Form3.Text2 = "" Then
            With RS4
                .Requery
                .Filter = "PROVEEDOR LIKE '*" & Form3.Text1 & "*'" & " AND FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
            With Form3
                With .DataGrid1
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 2500
                    .Columns(3).Width = 2000
                    .Columns(4).Width = 1500
                    .Columns(5).Width = 2500
                    .Columns(6).Width = 1500
                End With
            End With
        End If
        If Form3.Text1 = "" And Form3.Text2 = "" Then
            With RS4
                .Requery
                .Filter = "FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
           With Form3
                With .DataGrid1
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 2500
                    .Columns(3).Width = 2000
                    .Columns(4).Width = 1500
                    .Columns(5).Width = 2500
                    .Columns(6).Width = 1500
                End With
            End With
        End If
End Sub

Sub NCompra()
    On Error Resume Next
    With Form1
        For i = 0 To 5
            .Label1(i).Visible = True
        Next i
        .Text1(0).Visible = True
        .Text1(0) = ""
        For i = 0 To 4
            .Text2(i).Visible = True
            .Text2(i) = ""
        Next i
        .Command1.Visible = True
        .Command2.Visible = True
        .Command3.Visible = True
        .Command5.Visible = True
        '.Timer1.Enabled = True
        .Command5.SetFocus
    End With
End Sub

Sub OCompra()
    On Error Resume Next
    With Form1
        For i = 0 To 5
            .Label1(i).Visible = False
        Next i
        .Text1(0).Visible = False
        .Text1(0) = ""
        For i = 0 To 4
            .Text2(i).Visible = False
            .Text2(i) = ""
        Next i
        .Command1.Visible = False
        .Command2.Visible = False
        .Command3.Visible = False
        .Command5.Visible = False
        '.Timer1.Enabled = False
        If .MSComm1.PortOpen = True Then
            .MSComm1.PortOpen = False
        End If
    End With
End Sub
