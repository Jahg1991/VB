Attribute VB_Name = "Module1"
Global CN As New ADODB.Connection

Global RS1 As New ADODB.Recordset
Global RS2 As New ADODB.Recordset
Global RS3 As New ADODB.Recordset
Global RS4 As New ADODB.Recordset
Global RS5 As New ADODB.Recordset

Global IdProveedor As Integer
Global IdArticulo(1 To 15) As Integer
Global LArticulo As Integer

Global TPeso(1 To 15) As Double
Global TSubtotal(1 To 15) As Double
Global Precio(1 To 15) As Double

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

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

Sub BVentas()
    
    On Error Resume Next
        
        If Form3.Text1 <> "" Then
            
            With RS4
                .Requery
                .Filter = "PROVEEDOR LIKE '*" & Form3.Text1 & "*'" & " AND FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
            
            With RS5
                .Requery
                .Filter = "PROVEEDOR LIKE '*" & Form3.Text1 & "*'" & " AND FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
            
            With Form3
                
                With .DataGrid1
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 1500
                    .Columns(3).Width = 2500
                    .Columns(4).Width = 1500
                End With
            
            End With
        
        End If
        
        If Form3.Text1 = "" Then
            
            With RS4
                .Requery
                .Filter = "FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
            
            With RS5
                .Requery
                .Filter = "FECHA >= " & Form3.DTPicker1 & " AND FECHA <= " & Form3.DTPicker2
            End With
           
           With Form3
                
                With .DataGrid1
                    .Columns(0).Width = 1500
                    .Columns(1).Width = 2500
                    .Columns(2).Width = 1500
                    .Columns(3).Width = 2500
                    .Columns(4).Width = 1500
                End With
            
            End With
        
        End If

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
    
    With RS5
        
        If .State = 1 Then .Close
            
        .Open "Select * from d_compras", CN, adOpenStatic, adLockOptimistic
        .Requery
    
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

Sub CExportarExcelComprasDetalle()

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
    Dim l As Long
    Dim m As Long
    Dim art As Long
    Dim numFilas As Long
    
    i = 1
    j = 1
    
    l = 1
    m = 1
    
    numFilas = 0
    exportFileName = Form3.CommonDialog1.FileName
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    
    xlApp.DisplayAlerts = False
    
    For l = 1 To 1000
        
        For m = 1 To 6
            
            xlSheet.Cells(l, m).WrapText = True
        
        Next
    
    Next
    
    
    While (Not RS5.EOF)
        
        xlSheet.Cells(i, 1) = RS5.Fields(0).Name
        xlSheet.Cells(i, 2).Value = RS5(0)
        xlSheet.Cells(i, 3) = RS5.Fields(1).Name
        xlSheet.Cells(i, 4).Value = RS5(1)
        
        i = i + 1
        
        xlSheet.Cells(i, 1) = RS5.Fields(62).Name
        xlSheet.Cells(i, 2).Value = RS5(62)
        
        i = i + 2
        
        xlSheet.Cells(i, 2) = "Articulo"
        xlSheet.Cells(i, 3) = "Peso"
        xlSheet.Cells(i, 4) = "Precio"
        xlSheet.Cells(i, 5) = "Subtotal"
        
        i = i + 1
        
        'Articulo 1
        art = 1
        xlSheet.Cells(i, 1).Value = art
        xlSheet.Cells(i, 2).Value = RS5(2)
        xlSheet.Cells(i, 3).Value = RS5(3)
        xlSheet.Cells(i, 4).Value = RS5(4)
        xlSheet.Cells(i, 5).Value = RS5(5)
        
        i = i + 1
        art = art + 1
        
        'Articulo 2
        If RS5(6) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(6)
            xlSheet.Cells(i, 3).Value = RS5(7)
            xlSheet.Cells(i, 4).Value = RS5(8)
            xlSheet.Cells(i, 5).Value = RS5(9)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 3
        If RS5(10) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(10)
            xlSheet.Cells(i, 3).Value = RS5(11)
            xlSheet.Cells(i, 4).Value = RS5(12)
            xlSheet.Cells(i, 5).Value = RS5(13)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 4
        If RS5(14) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(14)
            xlSheet.Cells(i, 3).Value = RS5(15)
            xlSheet.Cells(i, 4).Value = RS5(16)
            xlSheet.Cells(i, 5).Value = RS5(17)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 5
        If RS5(18) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(18)
            xlSheet.Cells(i, 3).Value = RS5(19)
            xlSheet.Cells(i, 4).Value = RS5(20)
            xlSheet.Cells(i, 5).Value = RS5(21)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 6
        If RS5(22) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(22)
            xlSheet.Cells(i, 3).Value = RS5(23)
            xlSheet.Cells(i, 4).Value = RS5(24)
            xlSheet.Cells(i, 5).Value = RS5(25)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 7
        If RS5(26) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(26)
            xlSheet.Cells(i, 3).Value = RS5(27)
            xlSheet.Cells(i, 4).Value = RS5(28)
            xlSheet.Cells(i, 5).Value = RS5(29)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 8
        If RS5(30) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(30)
            xlSheet.Cells(i, 3).Value = RS5(31)
            xlSheet.Cells(i, 4).Value = RS5(32)
            xlSheet.Cells(i, 5).Value = RS5(33)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 9
        If RS5(34) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(34)
            xlSheet.Cells(i, 3).Value = RS5(35)
            xlSheet.Cells(i, 4).Value = RS5(36)
            xlSheet.Cells(i, 5).Value = RS5(37)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 10
        If RS5(38) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(38)
            xlSheet.Cells(i, 3).Value = RS5(39)
            xlSheet.Cells(i, 4).Value = RS5(40)
            xlSheet.Cells(i, 5).Value = RS5(41)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 11
        If RS5(42) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(42)
            xlSheet.Cells(i, 3).Value = RS5(43)
            xlSheet.Cells(i, 4).Value = RS5(44)
            xlSheet.Cells(i, 5).Value = RS5(45)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 12
        If RS5(46) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(46)
            xlSheet.Cells(i, 3).Value = RS5(47)
            xlSheet.Cells(i, 4).Value = RS5(48)
            xlSheet.Cells(i, 5).Value = RS5(49)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 13
        If RS5(50) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(50)
            xlSheet.Cells(i, 3).Value = RS5(51)
            xlSheet.Cells(i, 4).Value = RS5(52)
            xlSheet.Cells(i, 5).Value = RS5(53)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 14
        If RS5(54) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(54)
            xlSheet.Cells(i, 3).Value = RS5(55)
            xlSheet.Cells(i, 4).Value = RS5(56)
            xlSheet.Cells(i, 5).Value = RS5(57)
            
            i = i + 1
            art = art + 1
        
        End If
        
        'Articulo 15
        If RS5(58) <> "" Then
        
            xlSheet.Cells(i, 1).Value = art
            xlSheet.Cells(i, 2).Value = RS5(58)
            xlSheet.Cells(i, 3).Value = RS5(59)
            xlSheet.Cells(i, 4).Value = RS5(60)
            xlSheet.Cells(i, 5).Value = RS5(61)
            
            i = i + 1
            art = art + 1
        
        End If
        
        xlSheet.Cells(i, 2).Value = "Total"
        xlSheet.Cells(i, 3).Value = RS5(63)
        xlSheet.Cells(i, 5).Value = RS5(64)
        
        i = i + 3
        
        RS5.MoveNext
    
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
            .Columns(2).Width = 1500
            .Columns(3).Width = 2500
            .Columns(4).Width = 1500
        End With
        
        .DTPicker1.Value = Date
        .DTPicker2.Value = Date
    
    End With

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

Sub ICompra()
    
    On Error Resume Next
    
    With RS3
        
        .AddNew
            .Fields("FECHA") = Date
            
            .Fields("ID_PROVEEDOR") = IdProveedor
            .Fields("TPESO") = Form1.Text16(1)
            .Fields("TCOSTO") = Form1.Text16(2)
            .Fields("TICKET") = Form1.Text17
            
            .Fields("ID_ARTICULO1") = IdArticulo(1)
            .Fields("PRECIO1") = Form1.Text1(2)
            .Fields("PESO1") = Form1.Text1(3)
            .Fields("SUBTOTAL1") = Form1.Text1(4)
            
            If Form1.Text2(1) <> "" And Form1.Text2(2) <> "" And Form1.Text2(3) <> "" And Form1.Text2(4) <> "" Then
                .Fields("ID_ARTICULO2") = IdArticulo(2)
                .Fields("PRECIO2") = Form1.Text2(2)
                .Fields("PESO2") = Form1.Text2(3)
                .Fields("SUBTOTAL2") = Form1.Text2(4)
            End If
            
            If Form1.Text3(1) <> "" And Form1.Text3(2) <> "" And Form1.Text3(3) <> "" And Form1.Text3(4) <> "" Then
                .Fields("ID_ARTICULO3") = IdArticulo(3)
                .Fields("PRECIO3") = Form1.Text3(2)
                .Fields("PESO3") = Form1.Text3(3)
                .Fields("SUBTOTAL3") = Form1.Text3(4)
            End If
            
            If Form1.Text4(1) <> "" And Form1.Text4(2) <> "" And Form1.Text4(3) <> "" And Form1.Text4(4) <> "" Then
                .Fields("ID_ARTICULO4") = IdArticulo(4)
                .Fields("PRECIO4") = Form1.Text4(2)
                .Fields("PESO4") = Form1.Text4(3)
                .Fields("SUBTOTAL4") = Form1.Text4(4)
            End If
            
            If Form1.Text5(1) <> "" And Form1.Text5(2) <> "" And Form1.Text5(3) <> "" And Form1.Text5(4) <> "" Then
                .Fields("ID_ARTICULO5") = IdArticulo(5)
                .Fields("PRECIO5") = Form1.Text5(2)
                .Fields("PESO5") = Form1.Text5(3)
                .Fields("SUBTOTAL5") = Form1.Text5(4)
            End If
            
            If Form1.Text6(1) <> "" And Form1.Text6(2) <> "" And Form1.Text6(3) <> "" And Form1.Text6(4) <> "" Then
                .Fields("ID_ARTICULO6") = IdArticulo(6)
                .Fields("PRECIO6") = Form1.Text6(2)
                .Fields("PESO6") = Form1.Text6(3)
                .Fields("SUBTOTAL6") = Form1.Text6(4)
            End If
            
            If Form1.Text7(1) <> "" And Form1.Text7(2) <> "" And Form1.Text7(3) <> "" And Form1.Text7(4) <> "" Then
                .Fields("ID_ARTICULO7") = IdArticulo(7)
                .Fields("PRECIO7") = Form1.Text7(2)
                .Fields("PESO7") = Form1.Text7(3)
                .Fields("SUBTOTAL7") = Form1.Text7(4)
            End If
            
            If Form1.Text8(1) <> "" And Form1.Text8(2) <> "" And Form1.Text8(3) <> "" And Form1.Text8(4) <> "" Then
                .Fields("ID_ARTICULO8") = IdArticulo(8)
                .Fields("PRECIO8") = Form1.Text8(2)
                .Fields("PESO8") = Form1.Text8(3)
                .Fields("SUBTOTAL8") = Form1.Text8(4)
            End If
            
            If Form1.Text9(1) <> "" And Form1.Text9(2) <> "" And Form1.Text9(3) <> "" And Form1.Text9(4) <> "" Then
                .Fields("ID_ARTICULO9") = IdArticulo(9)
                .Fields("PRECIO9") = Form1.Text9(2)
                .Fields("PESO9") = Form1.Text9(3)
                .Fields("SUBTOTAL9") = Form1.Text9(4)
            End If
            
            If Form1.Text10(1) <> "" And Form1.Text10(2) <> "" And Form1.Text10(3) <> "" And Form1.Text10(4) <> "" Then
                .Fields("ID_ARTICULO10") = IdArticulo(10)
                .Fields("PRECIO10") = Form1.Text10(2)
                .Fields("PESO10") = Form1.Text10(3)
                .Fields("SUBTOTAL10") = Form1.Text10(4)
            End If
            
            If Form1.Text11(1) <> "" And Form1.Text11(2) <> "" And Form1.Text11(3) <> "" And Form1.Text11(4) <> "" Then
                .Fields("ID_ARTICULO11") = IdArticulo(11)
                .Fields("PRECIO11") = Form1.Text11(2)
                .Fields("PESO11") = Form1.Text11(3)
                .Fields("SUBTOTAL11") = Form1.Text11(4)
            End If
            
            If Form1.Text12(1) <> "" And Form1.Text12(2) <> "" And Form1.Text12(3) <> "" And Form1.Text12(4) <> "" Then
                .Fields("ID_ARTICULO12") = IdArticulo(12)
                .Fields("PRECIO12") = Form1.Text12(2)
                .Fields("PESO12") = Form1.Text12(3)
                .Fields("SUBTOTAL12") = Form1.Text12(4)
            End If
            
            If Form1.Text13(1) <> "" And Form1.Text13(2) <> "" And Form1.Text13(3) <> "" And Form1.Text13(4) <> "" Then
                .Fields("ID_ARTICULO13") = IdArticulo(13)
                .Fields("PRECIO13") = Form1.Text13(2)
                .Fields("PESO13") = Form1.Text13(3)
                .Fields("SUBTOTAL13") = Form1.Text13(4)
            End If
            
            If Form1.Text14(1) <> "" And Form1.Text14(2) <> "" And Form1.Text14(3) <> "" And Form1.Text14(4) <> "" Then
                .Fields("ID_ARTICULO14") = IdArticulo(14)
                .Fields("PRECIO14") = Form1.Text14(2)
                .Fields("PESO14") = Form1.Text14(3)
                .Fields("SUBTOTAL14") = Form1.Text14(4)
            End If
            
            If Form1.Text15(1) <> "" And Form1.Text15(2) <> "" And Form1.Text15(3) <> "" And Form1.Text15(4) <> "" Then
                .Fields("ID_ARTICULO15") = IdArticulo(15)
                .Fields("PRECIO15") = Form1.Text15(2)
                .Fields("PESO15") = Form1.Text15(3)
                .Fields("SUBTOTAL15") = Form1.Text15(4)
            End If
            
            .Update
        
            MsgBox "Compra guardada", vbOKOnly, "Terminado"
        
    End With
    
    IPrograma

End Sub

Sub IPrograma()
    
    IdProveedor = 0
    
    For i = 1 To 15
        TPeso(i) = 0
        IdArticulo(i) = 0
        TSubtotal(i) = 0
        Precio(i) = 0
    Next
    
    With Form1
        
        For i = 1 To 2
            .Text16(i) = 0
        Next
        
        .Command16.Enabled = False
        
        For i = 1 To 2
            .Command1(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text1(i).Enabled = False
            .Text1(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command2(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text2(i).Enabled = False
            .Text2(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command3(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text3(i).Enabled = False
            .Text3(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command4(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text4(i).Enabled = False
            .Text4(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command5(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text5(i).Enabled = False
            .Text5(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command6(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text6(i).Enabled = False
            .Text6(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command7(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text7(i).Enabled = False
            .Text7(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command8(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text8(i).Enabled = False
            .Text8(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command9(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text9(i).Enabled = False
            .Text9(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command10(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text10(i).Enabled = False
            .Text10(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command11(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text11(i).Enabled = False
            .Text11(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command12(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text12(i).Enabled = False
            .Text12(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command13(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text13(i).Enabled = False
            .Text13(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command14(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text14(i).Enabled = False
            .Text14(i).Text = ""
        Next
        
        For i = 1 To 2
            .Command15(i).Enabled = False
        Next
        
        For i = 1 To 4
            .Text15(i).Enabled = False
            .Text15(i).Text = ""
        Next
        
        .Text0.Text = ""
    
    End With

End Sub

Sub LCompra()
    
    On Error Resume Next
    
    With RS4
        .Requery
        .Filter = ""
        .MoveFirst
    End With
    
    With Form3
        
        Set .DataGrid1.DataSource = RS4
        
        .Text1 = ""
        .DTPicker1 = Date
        .DTPicker2 = Date
        
        With .DataGrid1
            .Columns(0).Width = 1500
            .Columns(1).Width = 2500
            .Columns(2).Width = 1500
            .Columns(3).Width = 2500
            .Columns(4).Width = 1500
        End With
    
    End With

End Sub

Function LeerPuertoBascula() As String

    'On Error GoTo err
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
    
    Form1.txtPeso.Text = LeerPuertoBascula

'err:
 '   LeerPuertoBascula = "ERROR"
  '  Form1.txtPeso.Text = LeerPuertoBascula

End Function

Sub Main()
    
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BD.mdb;Persist Security Info=False"
    End With
    
    CargarRS
    
    IPrograma
    
    Form1.Show

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
        
        .DataGrid1.Columns(1).Width = 8300
        RS1.MoveFirst
    
    End With

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
        
        .DataGrid1.Columns(1).Width = 11000
        RS2.MoveFirst
    End With

End Sub

Sub NCompra()
    
    On Error Resume Next
    
    With Form1
        .Frame1.Visible = True
    End With
    
    IPrograma
    
End Sub

Sub OCompra()
    
    On Error Resume Next
    
    With Form1
        
        .Frame1.Visible = False
        
        If .MSComm1.PortOpen = True Then
            .MSComm1.PortOpen = False
        End If
    
    End With

End Sub

Sub SBProveedor()
    
    On Error Resume Next
    
    If Form2.Text2.Visible = True Then
        
        With RS2
            IdProveedor = .Fields("ID")
            Form1.Text0.Text = .Fields("NOMBRE")
            Form1.Command1(1).Enabled = True
            Form1.Command1(1).SetFocus
        End With
    
    End If
    
    If Form2.Text1.Visible = True Then
        
        With RS1
            
            If LArticulo = 1 Then
                IdArticulo(1) = .Fields("ID")
                Precio(1) = .Fields("PRECIO")
                Form1.Text1(1).Text = .Fields("NOMBRE")
                Form1.Text1(2).Text = .Fields("PRECIO")
                Form1.Text1(2).Enabled = True
                Form1.Command1(2).Enabled = True
                Form1.Command16.Enabled = True
                Form1.Command1(2).SetFocus
            End If
            
            If LArticulo = 2 Then
                IdArticulo(2) = .Fields("ID")
                Precio(2) = .Fields("PRECIO")
                Form1.Text2(1).Text = .Fields("NOMBRE")
                Form1.Text2(2).Text = .Fields("PRECIO")
                Form1.Text2(2).Enabled = True
                Form1.Command2(2).Enabled = True
                Form1.Command2(2).SetFocus
            End If
            
            If LArticulo = 3 Then
                IdArticulo(3) = .Fields("ID")
                Precio(3) = .Fields("PRECIO")
                Form1.Text3(1).Text = .Fields("NOMBRE")
                Form1.Text3(2).Text = .Fields("PRECIO")
                Form1.Text3(2).Enabled = True
                Form1.Command3(2).Enabled = True
                Form1.Command3(2).SetFocus
            End If
            
            If LArticulo = 4 Then
                IdArticulo(4) = .Fields("ID")
                Precio(4) = .Fields("PRECIO")
                Form1.Text4(1).Text = .Fields("NOMBRE")
                Form1.Text4(2).Text = .Fields("PRECIO")
                Form1.Text4(2).Enabled = True
                Form1.Command4(2).Enabled = True
                Form1.Command4(2).SetFocus
            End If
            
            If LArticulo = 5 Then
                IdArticulo(5) = .Fields("ID")
                Precio(5) = .Fields("PRECIO")
                Form1.Text5(1).Text = .Fields("NOMBRE")
                Form1.Text5(2).Text = .Fields("PRECIO")
                Form1.Text5(2).Enabled = True
                Form1.Command5(2).Enabled = True
                Form1.Command5(2).SetFocus
            End If
            
            If LArticulo = 6 Then
                IdArticulo(6) = .Fields("ID")
                Precio(6) = .Fields("PRECIO")
                Form1.Text6(1).Text = .Fields("NOMBRE")
                Form1.Text6(2).Text = .Fields("PRECIO")
                Form1.Text6(2).Enabled = True
                Form1.Command6(2).Enabled = True
                Form1.Command6(2).SetFocus
            End If
            
            If LArticulo = 7 Then
                IdArticulo(7) = .Fields("ID")
                Precio(7) = .Fields("PRECIO")
                Form1.Text7(1).Text = .Fields("NOMBRE")
                Form1.Text7(2).Text = .Fields("PRECIO")
                Form1.Text7(2).Enabled = True
                Form1.Command7(2).Enabled = True
                Form1.Command7(2).SetFocus
            End If
            
            If LArticulo = 8 Then
                IdArticulo(8) = .Fields("ID")
                Precio(8) = .Fields("PRECIO")
                Form1.Text8(1).Text = .Fields("NOMBRE")
                Form1.Text8(2).Text = .Fields("PRECIO")
                Form1.Text8(2).Enabled = True
                Form1.Command8(2).Enabled = True
                Form1.Command8(2).SetFocus
            End If
            
            If LArticulo = 9 Then
                IdArticulo(9) = .Fields("ID")
                Precio(9) = .Fields("PRECIO")
                Form1.Text9(1).Text = .Fields("NOMBRE")
                Form1.Text9(2).Text = .Fields("PRECIO")
                Form1.Text9(2).Enabled = True
                Form1.Command9(2).Enabled = True
                Form1.Command9(2).SetFocus
            End If
            
            If LArticulo = 10 Then
                IdArticulo(10) = .Fields("ID")
                Precio(10) = .Fields("PRECIO")
                Form1.Text10(1).Text = .Fields("NOMBRE")
                Form1.Text10(2).Text = .Fields("PRECIO")
                Form1.Text10(2).Enabled = True
                Form1.Command10(2).Enabled = True
                Form1.Command10(2).SetFocus
            End If
            
            If LArticulo = 11 Then
                IdArticulo(11) = .Fields("ID")
                Precio(11) = .Fields("PRECIO")
                Form1.Text11(1).Text = .Fields("NOMBRE")
                Form1.Text11(2).Text = .Fields("PRECIO")
                Form1.Text11(2).Enabled = True
                Form1.Command11(2).Enabled = True
                Form1.Command11(2).SetFocus
            End If
            
            If LArticulo = 12 Then
                IdArticulo(12) = .Fields("ID")
                Precio(12) = .Fields("PRECIO")
                Form1.Text12(1).Text = .Fields("NOMBRE")
                Form1.Text12(2).Text = .Fields("PRECIO")
                Form1.Text12(2).Enabled = True
                Form1.Command12(2).Enabled = True
                Form1.Command12(2).SetFocus
            End If
            
            If LArticulo = 13 Then
                IdArticulo(13) = .Fields("ID")
                Precio(13) = .Fields("PRECIO")
                Form1.Text13(1).Text = .Fields("NOMBRE")
                Form1.Text13(2).Text = .Fields("PRECIO")
                Form1.Text13(2).Enabled = True
                Form1.Command13(2).Enabled = True
                Form1.Command13(2).SetFocus
            End If
            
            If LArticulo = 14 Then
                IdArticulo(14) = .Fields("ID")
                Precio(14) = .Fields("PRECIO")
                Form1.Text14(1).Text = .Fields("NOMBRE")
                Form1.Text14(2).Text = .Fields("PRECIO")
                Form1.Text14(2).Enabled = True
                Form1.Command14(2).Enabled = True
                Form1.Command14(2).SetFocus
            End If
            
            If LArticulo = 15 Then
                IdArticulo(15) = .Fields("ID")
                Precio(15) = .Fields("PRECIO")
                Form1.Text15(1).Text = .Fields("NOMBRE")
                Form1.Text15(2).Text = .Fields("PRECIO")
                Form1.Text15(2).Enabled = True
                Form1.Command15(2).Enabled = True
                Form1.Command15(2).SetFocus
            End If
        
        End With
    
    End If
    
    Unload Form2

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
