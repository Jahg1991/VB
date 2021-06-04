Attribute VB_Name = "Module1"
Global Cn As New ADODB.Connection

Global RsEmpresaAbierta As New ADODB.Recordset
Global RsEmpresa As New ADODB.Recordset

Global GTipo As Integer
    '1  Añadir Empresa
    '2  Editar Empresa

Global CTipo As Integer
    '1  Form1
    '2  Form2 Añadir Empresa
    '3  Form3 Editar Empresa
    
Global EmpAbierta As Integer

Sub OpenBd()

    With Cn
    
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\JAHG Software\Admin Clintes\BD.mdb;Persist Security Info=False"
        
    End With

End Sub

Sub CargarRs()
    
    With RsEmpresaAbierta
        
        If .State = 1 Then .Close
        
            .Open "Select id,nombre from Empresas", Cn, adOpenStatic, adLockOptimistic
            .Filter = ""
            .Requery
        
    End With
    
    With RsEmpresa
        
        If .State = 1 Then .Close
        
            .Open "Select * from Empresas", Cn, adOpenStatic, adLockOptimistic
            .Filter = ""
            .Requery
        
    End With
    
End Sub

Sub InitForm()

    If CTipo = 1 Then

        CargarRs
        
        With Form1
        
            .List1.Clear
            
            Do Until RsEmpresaAbierta.EOF
            
                .List1.AddItem RsEmpresaAbierta("Nombre")
                .List1.Text = RsEmpresaAbierta("Nombre")
                
                RsEmpresaAbierta.MoveNext
                
            Loop
            
            .Show
            
            .List1.SetFocus
            
        
        End With
    
    End If
    
    If CTipo = 2 Then
    
        With Form2
        
            .Caption = "Crear Empresa"
            
            For i = 0 To 16
        
                .Text1(i).Text = ""
            
            Next i
            
            .Show
            
            .SSTab1.Tab = 0
            
            .Text1(0).SetFocus
            
        End With
        
    End If
    
    If CTipo = 3 Then
    
        With Form2
            
            .Caption = "Editar Empresa"
            
            With RsEmpresa
                
                .Filter = "Id = '" & EmpAbierta & "'"
            
            End With
            
            For i = 0 To 16

                Set .Text1(i).DataSource = RsEmpresa
            
            Next i
            
            .Text1(0).DataField = ("Nombre")
            .Text1(1).DataField = ("RFC")
            .Text1(2).DataField = ("NSS")
            .Text1(3).DataField = ("INE")
            .Text1(4).DataField = ("CURP")
            .Text1(5).DataField = ("Telefono")
            .Text1(6).DataField = ("Celular")
            .Text1(7).DataField = ("Correo")
            .Text1(8).DataField = ("Calle")
            .Text1(9).DataField = ("Numero int")
            .Text1(10).DataField = ("Numero ext")
            .Text1(11).DataField = ("Colonia")
            .Text1(12).DataField = ("CP")
            .Text1(13).DataField = ("Localidad")
            .Text1(14).DataField = ("Ciudad")
            .Text1(15).DataField = ("Estado")
            .Text1(16).DataField = ("Pais")
            
            .Show
            
            RsEmpresa.MoveFirst
            
            .SSTab1.Tab = 0
            
            .Text1(0).SetFocus
            
        End With
        
    End If

End Sub

Sub Guardar()

    If GTipo = 1 Then
    
        If Form2.Text1(0) <> "" Then
    
            With RsEmpresa
            
                .AddNew
                    .Fields("Nombre") = Form2.Text1(0)
                    .Fields("RFC") = Form2.Text1(1)
                    .Fields("NSS") = Form2.Text1(2)
                    .Fields("INE") = Form2.Text1(3)
                    .Fields("CURP") = Form2.Text1(4)
                    .Fields("Telefono") = Form2.Text1(5)
                    .Fields("Celular") = Form2.Text1(6)
                    .Fields("Correo") = Form2.Text1(7)
                    .Fields("Calle") = Form2.Text1(8)
                    .Fields("Numero int") = Form2.Text1(9)
                    .Fields("Numero ext") = Form2.Text1(10)
                    .Fields("Colonia") = Form2.Text1(11)
                    If Form2.Text1(12) = "" Then
                        Form2.Text1(12) = 47000
                        .Fields("CP") = Form2.Text1(12)
                    End If
                    .Fields("Localidad") = Form2.Text1(13)
                    .Fields("Ciudad") = Form2.Text1(14)
                    .Fields("Estado") = Form2.Text1(15)
                    .Fields("Pais") = Form2.Text1(16)
                .Update
            
            End With
            
            MsgBox "Listo!!!", vbOKOnly, "Terminado"
        
        Else
            
            MsgBox "El nombre es necesario", vbOKOnly, "Error"
        
        End If
    
    End If
    
    If GTipo = 2 Then
    
        If Form2.Text1(0) <> "" Then
                
            With RsEmpresa
                
                If Form2.Text1(12) = "" Then
                    Form2.Text1(12) = 47000
                End If
                
                .Update
                
            End With
                
            MsgBox "Listo!!!", vbOKOnly, "Terminado"
        
        Else
            
            MsgBox "El nombre es necesario", vbOKOnly, "Error"
        
        End If
    
    End If

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
    
    CTipo = 1
    
    InitForm
    
End Sub
