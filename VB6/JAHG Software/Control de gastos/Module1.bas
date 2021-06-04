Attribute VB_Name = "Module1"
Global Cn As New ADODB.Connection

Global RsGastos As New ADODB.Recordset
Global RsIngresos As New ADODB.Recordset
Global RsPresupuesto As New ADODB.Recordset
Global RsRepIngresosSemana As New ADODB.Recordset
Global RsRepIngresosMes As New ADODB.Recordset
Global RsRepGastosSemana As New ADODB.Recordset
Global RsRepGastosMes As New ADODB.Recordset

Global TipoReporte As Integer

Sub OpenBd()

    With Cn
    
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BD.mdb;Persist Security Info=False"
        
    End With

End Sub

Sub CargarRs()
    
    With RsGastos
        
        If .State = 1 Then .Close
        
            .Open "Select * from gastos", Cn, adOpenStatic, adLockOptimistic
            .Requery
        
    End With
    
    With RsIngresos
        
        If .State = 1 Then .Close
            
            .Open "Select * from ingresos", Cn, adOpenStatic, adLockOptimistic
            .Requery
    
    End With
    
    With RsPresupuesto
        
        If .State = 1 Then .Close
            
            .Open "Select presupuesto from presupuesto", Cn, adOpenStatic, adLockOptimistic
            .Requery
    
    End With
    
    With RsRepIngresosSemana
        
        If .State = 1 Then .Close
            
            .Open "Select * from ingresos_semana", Cn, adOpenStatic, adLockOptimistic
            .Requery
    
    End With
    
    With RsRepIngresosMes
        
        If .State = 1 Then .Close
            
            .Open "Select * from ingresos_mes", Cn, adOpenStatic, adLockOptimistic
            .Requery
    
    End With
    
    With RsRepGastosSemana
        
        If .State = 1 Then .Close
            
            .Open "Select * from gastos_semana", Cn, adOpenStatic, adLockOptimistic
            .Requery
    
    End With
    
    With RsRepGastosMes
        
        If .State = 1 Then .Close
            
            .Open "Select * from gastos_mes", Cn, adOpenStatic, adLockOptimistic
            .Requery
    
    End With
    
End Sub

Sub InitForm()

    CargarRs
    
    With Form1
        
        .Label1(16) = RsPresupuesto.Fields("presupuesto")
    
        .DTPicker1.Value = Date
        
        For i = 0 To 13
        
            .Text1(i).Text = 0
        
        Next i
    
    End With

End Sub

Sub guardar()

    With RsGastos
    
        .AddNew
            .Fields("Fecha") = Form1.DTPicker1
            .Fields("Casa") = Form1.Text1(0)
            .Fields("Carro") = Form1.Text1(1)
            .Fields("Hogar") = Form1.Text1(2)
            .Fields("Alfredo") = Form1.Text1(3)
            .Fields("Yisel") = Form1.Text1(4)
            .Fields("Mandado") = Form1.Text1(5)
            .Fields("Inbursa") = Form1.Text1(6)
            .Fields("Otros") = Form1.Text1(7)
            .Fields("Renta") = Form1.Text1(8)
            .Fields("Salidas") = Form1.Text1(9)
            .Fields("Viajes") = Form1.Text1(10)
        .Update
    
    End With
    
    With RsIngresos
    
        .AddNew
            .Fields("Fecha") = Form1.DTPicker1
            .Fields("Alfredo") = Form1.Text1(11)
            .Fields("Yisel") = Form1.Text1(12)
            .Fields("Inbursa") = Form1.Text1(13)
        .Update
        
    End With

End Sub

Sub CExportarExcel()

    On Error GoTo err

    Form1.CommonDialog1.DialogTitle = "Guardar como"
    Form1.CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    Form1.CommonDialog1.ShowSave
    
    If Form1.CommonDialog1.FileName = "" Then
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
    exportFileName = Form1.CommonDialog1.FileName
    
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    
    xlApp.DisplayAlerts = False
    
    If TipoReporte = 1 Then
        
        While (numFilas < RsRepIngresosSemana.Fields.Count)
            xlSheet.Cells(i, j) = RsRepIngresosSemana.Fields(numFilas).Name
            j = j + 1
            numFilas = numFilas + 1
        Wend
            
        i = i + 1
        
        While (Not RsRepIngresosSemana.EOF)
            j = 0
            
            While (j < numFilas)
                xlSheet.Cells(i, j + 1).Value = RsRepIngresosSemana(j)
                j = j + 1
            Wend
            
            i = i + 1
            RsRepIngresosSemana.MoveNext
        
        Wend
    
    End If
    
    If TipoReporte = 2 Then
        
        While (numFilas < RsRepIngresosMes.Fields.Count)
            xlSheet.Cells(i, j) = RsRepIngresosMes.Fields(numFilas).Name
            j = j + 1
            numFilas = numFilas + 1
        Wend
            
        i = i + 1
        
        While (Not RsRepIngresosMes.EOF)
            j = 0
            
            While (j < numFilas)
                xlSheet.Cells(i, j + 1).Value = RsRepIngresosMes(j)
                j = j + 1
            Wend
            
            i = i + 1
            RsRepIngresosMes.MoveNext
        
        Wend
    
    End If
    
    If TipoReporte = 3 Then
        
        While (numFilas < RsRepGastosSemana.Fields.Count)
            xlSheet.Cells(i, j) = RsRepGastosSemana.Fields(numFilas).Name
            j = j + 1
            numFilas = numFilas + 1
        Wend
            
        i = i + 1
        
        While (Not RsRepGastosSemana.EOF)
            j = 0
            
            While (j < numFilas)
                xlSheet.Cells(i, j + 1).Value = RsRepGastosSemana(j)
                j = j + 1
            Wend
            
            i = i + 1
            RsRepGastosSemana.MoveNext
        
        Wend
    
    End If
    
    If TipoReporte = 4 Then
        
        While (numFilas < RsRepGastosMes.Fields.Count)
            xlSheet.Cells(i, j) = RsRepGastosMes.Fields(numFilas).Name
            j = j + 1
            numFilas = numFilas + 1
        Wend
            
        i = i + 1
        
        While (Not RsRepGastosMes.EOF)
            j = 0
            
            While (j < numFilas)
                xlSheet.Cells(i, j + 1).Value = RsRepGastosMes(j)
                j = j + 1
            Wend
            
            i = i + 1
            RsRepGastosMes.MoveNext
        
        Wend
    
    End If
    
    xlSheet.SaveAs exportFileName
    xlBook.Close
    xlApp.Quit
    
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    MsgBox "Archivo creado correctamente", vbOKOnly, "Terminado"
    
    Form1.CommonDialog1.FileName = ""
    
    Exit Sub

err:
        
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        
        MsgBox "El archivo no ha podido ser creado" & vbNewLine & err.Description, vbOKOnly, "Error"

End Sub

Sub main()

    OpenBd
    
    InitForm
    
    Form1.Show
    
End Sub
