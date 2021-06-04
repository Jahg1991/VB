Attribute VB_Name = "Module1"
'***********************************
'Programa para el control del ganado
'Hecho por Juan Alfredo Hernández
'v1.0 08/03/2017
'***********************************
'Declaramos variables
Global VarFecha As Date
Global VarFondo As String
Global VarForm1State As Integer '0 Cerrada 1 Abierta
Global VarForm2State As Integer '0 Cerrada 1 Abierta
Global VarForm3State As Integer '0 Cerrada 1 Abierta
Global VarForm4State As Integer '0 Cerrada 1 Abierta
Global VarForm5State As Integer '0 Cerrada 1 Abierta
Global VarForm6State As Integer '0 Cerrada 1 Abierta
Global VarHato As Integer
Global VarProductor As Integer
Global VarTipo As String
Global VarTipoConexionProductor As Integer '1 Productor     2 Hato     3 Personal
'Declaramos las conexiones
Global BdConexion As New ADODB.Connection
'Declaramos los juegos de registros
Global BdRecordSet01 As New ADODB.Recordset
Global BdRecordSet02 As New ADODB.Recordset
Sub main()
    ProcAbrirConexion
    'Inicializamos variables
    VarProductor = 0
    VarHato = 0
    VarTipoConexionProductor = 0
    VarForm1State = 0
    VarForm2State = 0
    VarForm3State = 0
    VarForm4State = 0
    VarForm5State = 0
    VarForm6State = 0
    'Abrimos el formulario principal
    With Form1
        .Show
    End With
    VarForm1State = 1
    'Cargamos la imagen de fondo en el Picture
    ProcCargarFondo
End Sub
Sub ProcAbrirConexion()
    'Abrimos la conexión a la base de datos
    With BdConexion
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BD.mdb;Persist Security Info=False"
    End With
End Sub
Sub ProcActualizarFondo()
    'Abrimos la conexión
    With Form3
        With BdRecordSet01
            'Abrimos el juego de registros
            If .State = 1 Then .Close
            .Open "select * from Fondo", BdConexion, adOpenStatic, adLockOptimistic
            .Requery
        End With
        With .DataGrid1
            Set .DataSource = BdRecordSet01
            'Actualizamos los campos con los valores de las variables
            .Columns(1) = VarFondo
            .Columns(2) = VarFecha
        End With
        With BdRecordSet01
            .Update
            'Cerramos la conexión
            .Close
        End With
    End With
    VarForm3State = 0
    Unload Form3
End Sub
Sub ProcActualizarProductores()
    'Actualizamos el registro
    actualizar = MsgBox("¿Desea actualizar el registro?", vbExclamation + vbYesNo, "Confirmación de actualización")
        If actualizar = vbYes Then
            With BdRecordSet02
                .Update
            End With
            ProcHabilitarBusqueda
            msg = MsgBox("Registro actualizado", vbOKOnly, "Terminado")
        End If
End Sub
Sub ProcActualizacionTablas()
    With Form5
        With .Combo1
            If .Text = "Categoría" Then
                With BdRecordSet01
                    'Àbrimos la conexiòn a la base de datos
                    If .State = 1 Then .Close
                    .Open "select * from Categoria", BdConexion, adOpenStatic, adLockOptimistic
                    .Requery
                End With
            Else
                If .Text = "Causa de aborto" Then
                     With BdRecordSet01
                        'Àbrimos la conexiòn a la base de datos
                        If .State = 1 Then .Close
                        .Open "select * from CausaAborto", BdConexion, adOpenStatic, adLockOptimistic
                        .Requery
                    End With
                Else
                    If .Text = "Causa de no inseminar" Then
                            With BdRecordSet01
                                'Àbrimos la conexiòn a la base de datos
                                If .State = 1 Then .Close
                                .Open "select * from CausaNoInseminar", BdConexion, adOpenStatic, adLockOptimistic
                                .Requery
                            End With
                    Else
                        If .Text = "Causa de rechazo" Then
                            With BdRecordSet01
                                'Àbrimos la conexiòn a la base de datos
                                If .State = 1 Then .Close
                                .Open "select * from CausaRechazo", BdConexion, adOpenStatic, adLockOptimistic
                                .Requery
                            End With
                        Else
                            If .Text = "Diagnóstico de útero" Then
                                With BdRecordSet01
                                    'Àbrimos la conexiòn a la base de datos
                                    If .State = 1 Then .Close
                                    .Open "select * from DiagnosticoUtero", BdConexion, adOpenStatic, adLockOptimistic
                                    .Requery
                                End With
                            Else
                                If .Text = "Enfermedad de ovario" Then
                                    With BdRecordSet01
                                        'Àbrimos la conexiòn a la base de datos
                                        If .State = 1 Then .Close
                                        .Open "select * from EnfermedadOvario", BdConexion, adOpenStatic, adLockOptimistic
                                        .Requery
                                    End With
                                Else
                                    If .Text = "Enfermedad del útero" Then
                                        With BdRecordSet01
                                            'Àbrimos la conexiòn a la base de datos
                                            If .State = 1 Then .Close
                                            .Open "select * from EnfermedadUtero", BdConexion, adOpenStatic, adLockOptimistic
                                            .Requery
                                        End With
                                    Else
                                        If .Text = "Especificaciones de la muerte" Then
                                            With BdRecordSet01
                                                'Àbrimos la conexiòn a la base de datos
                                                If .State = 1 Then .Close
                                                .Open "select * from EspecificacionesMuerte", BdConexion, adOpenStatic, adLockOptimistic
                                                .Requery
                                            End With
                                        Else
                                            If .Text = "Especificaciones de venta" Then
                                                With BdRecordSet01
                                                    'Àbrimos la conexiòn a la base de datos
                                                    If .State = 1 Then .Close
                                                    .Open "select * from EspecificacionesVenta", BdConexion, adOpenStatic, adLockOptimistic
                                                    .Requery
                                                End With
                                            Else
                                                If .Text = "Estado de la cría" Then
                                                    With BdRecordSet01
                                                        'Àbrimos la conexiòn a la base de datos
                                                        If .State = 1 Then .Close
                                                        .Open "select * from EstadoCria", BdConexion, adOpenStatic, adLockOptimistic
                                                        .Requery
                                                    End With
                                                Else
                                                    If .Text = "Medicamentos de cuartos mamarios" Then
                                                       With BdRecordSet01
                                                            'Àbrimos la conexiòn a la base de datos
                                                            If .State = 1 Then .Close
                                                            .Open "select * from MedicacionCuartosMamarios", BdConexion, adOpenStatic, adLockOptimistic
                                                            .Requery
                                                       End With
                                                    Else
                                                        If .Text = "Medicación Genital" Then
                                                            With BdRecordSet01
                                                                'Àbrimos la conexiòn a la base de datos
                                                                If .State = 1 Then .Close
                                                                .Open "select * from MedicacionGenital", BdConexion, adOpenStatic, adLockOptimistic
                                                                .Requery
                                                            End With
                                                        Else
                                                            If .Text = "Medicamentos" Then
                                                                With BdRecordSet01
                                                                    'Àbrimos la conexiòn a la base de datos
                                                                    If .State = 1 Then .Close
                                                                    .Open "select * from Medicamento", BdConexion, adOpenStatic, adLockOptimistic
                                                                    .Requery
                                                                End With
                                                            Else
                                                                If .Text = "Raza" Then
                                                                    With BdRecordSet01
                                                                        'Àbrimos la conexiòn a la base de datos
                                                                        If .State = 1 Then .Close
                                                                        .Open "select * from Raza", BdConexion, adOpenStatic, adLockOptimistic
                                                                        .Requery
                                                                    End With
                                                                Else
                                                                    If .Text = "Resultado análisis" Then
                                                                        With BdRecordSet01
                                                                            'Àbrimos la conexiòn a la base de datos
                                                                            If .State = 1 Then .Close
                                                                            .Open "select * from ResultadoAnalisis", BdConexion, adOpenStatic, adLockOptimistic
                                                                            .Requery
                                                                        End With
                                                                    Else
                                                                        If .Text = "Sexo cría" Then
                                                                            With BdRecordSet01
                                                                                'Àbrimos la conexiòn a la base de datos
                                                                                If .State = 1 Then .Close
                                                                                .Open "select * from SexoCria", BdConexion, adOpenStatic, adLockOptimistic
                                                                                .Requery
                                                                            End With
                                                                        Else
                                                                            If .Text = "Tipo de análisis" Then
                                                                                With BdRecordSet01
                                                                                    'Àbrimos la conexiòn a la base de datos
                                                                                    If .State = 1 Then .Close
                                                                                    .Open "select * from TipoAnalisis", BdConexion, adOpenStatic, adLockOptimistic
                                                                                    .Requery
                                                                                End With
                                                                            Else
                                                                                If .Text = "Tipo de enfermedad" Then
                                                                                    With BdRecordSet01
                                                                                        'Àbrimos la conexiòn a la base de datos
                                                                                        If .State = 1 Then .Close
                                                                                        .Open "select * from TipoEnfermedad", BdConexion, adOpenStatic, adLockOptimistic
                                                                                        .Requery
                                                                                    End With
                                                                                Else
                                                                                    If .Text = "Tipo de parto" Then
                                                                                        With BdRecordSet01
                                                                                            'Àbrimos la conexiòn a la base de datos
                                                                                            If .State = 1 Then .Close
                                                                                            .Open "select * from TipoParto", BdConexion, adOpenStatic, adLockOptimistic
                                                                                            .Requery
                                                                                        End With
                                                                                    Else
                                                                                        If .Text = "Via de Aplicación" Then
                                                                                           With BdRecordSet01
                                                                                                'Àbrimos la conexiòn a la base de datos
                                                                                                If .State = 1 Then .Close
                                                                                                .Open "select * from ViaAplicacion", BdConexion, adOpenStatic, adLockOptimistic
                                                                                                .Requery
                                                                                            End With
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
        .DataGrid1.Caption = .Combo1.Text
        With .DataGrid1
            'Asignamos la conexion al grid
            Set .DataSource = BdRecordSet01
        End With
        ProcResizeGrid
    End With
End Sub
Sub ProcAsignaciones()
    If VarTipo = "Productores" Then
    'Ponemos el nombre del productor en un label
        With Form1
        'Condicional para productores
            With .Label2
                .Caption = "Productor: " + Form2.DataGrid1.Columns(1).Text
            End With
        End With
        With Form2
            With .DataGrid1
                'Ponemos el id del productor en un label
                VarProductor = .Columns(0).Value
            End With
        End With
    Else
        'Condicional para Hatos
        If VarTipo = "Hatos" Then
            'Ponemos el nombre del hato en un label
            With Form1
                With .Label3
                    .Caption = "Hato: " + Form2.DataGrid1.Columns(1).Text
                End With
                'Habilitamos Animales
                With .Animales
                    .Enabled = True
                End With
            End With
            'Ponemos el id del hato en un label
            With Form2
                With .DataGrid1
                    VarHato = .Columns(0).Value
                End With
            End With
        End If
    End If
    'Cerramos la conexiòn
    With BdRecordSet01
        .Close
    End With
    VarForm2State = 0
    Unload Form2
End Sub
Sub ProcBusquedaProductores()
    On Error Resume Next
    'Filtramos los registros segun el contenido del textbox
    With BdRecordSet02
        .Requery
        If option1.Value = True Then
            .Filter = "Nombre like '*" & Form4.Text9 & "*'"
        Else
            .Filter = ""
            .MoveFirst
        End If
    End With
    With Form4
        With .DataGrid3
            Set .DataSource = BdRecordSet02
        End With
    End With
End Sub
Sub ProcProductoresGridColumnas()
    With Form4
        If VarTipoConexionProductor = 1 Then 'Productor
            With .DataGrid3
                .Columns(2).Visible = False
                .Columns(3).Visible = False
                .Columns(4).Visible = False
            End With
            ProcResizeGrid
        Else
            If VarTipoConexionProductor = 2 Then 'Hatos
                With .DataGrid3
                    .Columns(1).Visible = False
                    .Columns(3).Visible = False
                    .Columns(4).Visible = False
                    .Columns(5).Visible = False
                End With
                ProcResizeGrid
            Else
                If VarTipoConexionProductor = 3 Then 'Personal
                    With .DataGrid3
                        .Columns(1).Visible = False
                        .Columns(3).Visible = False
                        .Columns(4).Visible = False
                        .Columns(5).Visible = False
                    End With
                    ProcResizeGrid
                End If
            End If
        End If
    End With
End Sub
Sub ProcCargarFondo()
    On Error Resume Next
    'Abrimos la conexión
    With BdRecordSet01
        If .State = 1 Then .Close
        .Open "select * from Fondo", BdConexion, adOpenStatic, adLockOptimistic
        .Requery
    End With
    With Form3
        With .DataGrid1
            Set .DataSource = BdRecordSet01
            VarFondo = .Columns(1).Text
        End With
    End With
    With Form1
        With .Image1
            'Cambiamos la imagen de fondo
            .Picture = LoadPicture(VarFondo)
        End With
    End With
    'Cerramos la conexión
    With BdRecordSet01
        .Close
    End With
    VarForm3State = 0
    Unload Form3
End Sub
Sub ProcCerrarConexion()
    'Cerramos las conexiones
    With BdRecordSet01
        If .State = 1 Then .Close
    End With
    With BdRecordSet02
        If .State = 1 Then .Close
    End With
    With BdConexion
        If .State = 1 Then .Close
    End With
End Sub
Sub ProcCerrarProductor()
    VarForm4State = 0
    With Form1
        'Deshabilitamos animales
        With .Animales
            .Enabled = False
        End With
        'Limpiamos etiquetas
        With .Label2
            .Caption = ""
        End With
        With .Label3
            .Caption = ""
        End With
        'Inicializamos variables
        VarProductor = 0
        VarHato = 0
        'Inabilitamos el seleccionar hatos
        With .SeleccionarHato
            .Enabled = False
        End With
    End With
End Sub
Sub ProcConexionAnimales()
    VarForm6State = 1
    With Form6
        'Mostramos el Formulario
        .Show
        'Abrimos la conexion
        With BdRecordSet01
            If .State = 1 Then .Close
            .Open "select * from VAnimales where Productor = " & VarProductor & "and Hato = " & VarHato & "and Estado = 0", BdConexion, adOpenStatic, adLockOptimistic  '0 Normal 1 Muerta 2 Vendida
            .Requery
        End With
        'Asignamos los datos al Grid
        With .DataGrid1
            Set .DataSource = BdRecordSet01
            .Caption = "General de animales"
            'Ocultamos columnas
            .Columns(0).Visible = False
            .Columns(1).Visible = False
            .Columns(9).Visible = False
        End With
    End With
End Sub
Sub ProcConexionNuevoProductor()
    ProcLimpiarFormulario
    With Form4
        If VarTipoConexionProductor = 1 Then 'Productor
            .Caption = "Productores"
            'Abrimos las conexiones
            With BdRecordSet01
                If .State = 1 Then .Close
                .Open "select * from Productores", BdConexion, adOpenStatic, adLockOptimistic
                .Requery
            End With
            With BdRecordSet02
                If .State = 1 Then .Close
                .Open "select * from Productores", BdConexion, adOpenStatic, adLockOptimistic
                .Requery
            End With
            'Asignamos el juego de registros al grid
            With .DataGrid1
                Set .DataSource = BdRecordSet01
            End With
            With .DataGrid2
                Set .DataSource = BdRecordSet02
            End With
            With .DataGrid3
                Set .DataSource = BdRecordSet02
            End With
            'Asignamos los campos a los textbox
            With .Text4
                Set .DataSource = BdRecordSet02
                .DataField = "Nombre"
            End With
            With .Text5
                Set .DataSource = BdRecordSet02
                .DataField = "Direccion"
            End With
            With .Text6
                Set .DataSource = BdRecordSet02
                .DataField = "Ciudad"
            End With
        Else
            If VarTipoConexionProductor = 2 Then 'Hatos
                .Caption = "Hatos de " + Form1.Label2.Caption
                'Abrimos las conexiones
                With BdRecordSet01
                    If .State = 1 Then .Close
                    .Open "select * from Hatos where productor = " & VarProductor, BdConexion, adOpenStatic, adLockOptimistic
                    .Requery
                End With
                With BdRecordSet02
                    If .State = 1 Then .Close
                    .Open "select * from Hatos where productor = " & VarProductor, BdConexion, adOpenStatic, adLockOptimistic
                    .Requery
                End With
                'Asignamos el juego de registros al grid
                With .DataGrid1
                    Set .DataSource = BdRecordSet01
                End With
                With .DataGrid2
                    Set .DataSource = BdRecordSet02
                End With
                With .DataGrid3
                    Set .DataSource = BdRecordSet02
                End With
                'Asignamos los campos a los textbox
                With .Text4
                    Set .DataSource = BdRecordSet02
                    .DataField = "Nombre"
                End With
                With .Text5
                    Set .DataSource = BdRecordSet02
                    .DataField = "Direccion"
                End With
                With .Text6
                    Set .DataSource = BdRecordSet02
                    .DataField = "Ciudad"
                End With
            Else
                If VarTipoConexionProductor = 3 Then 'Personal
                    .Caption = "Personal de " + Form1.Label2.Caption
                    'Abrimos las conexiones
                    With BdRecordSet01
                        If .State = 1 Then .Close
                        .Open "select * from Personal where productor = " & VarProductor, BdConexion, adOpenStatic, adLockOptimistic
                        .Requery
                    End With
                    With BdRecordSet02
                        If .State = 1 Then .Close
                        .Open "select * from Personal where productor = " & VarProductor, BdConexion, adOpenStatic, adLockOptimistic
                        .Requery
                    End With
                    'Asignamos el juego de registros al grid
                    With .DataGrid1
                        Set .DataSource = BdRecordSet01
                    End With
                    With .DataGrid2
                        Set .DataSource = BdRecordSet02
                    End With
                    With .DataGrid3
                        Set .DataSource = BdRecordSet02
                    End With
                    'Asignamos los campos a los textbox
                    With .Text4
                        Set .DataSource = BdRecordSet02
                        .DataField = "Nombre"
                    End With
                    With .Text5
                        Set .DataSource = BdRecordSet02
                        .DataField = "Direccion"
                    End With
                    With .Text6
                        Set .DataSource = BdRecordSet02
                        .DataField = "Ciudad"
                    End With
                End If
            End If
        End If
    End With
End Sub
Sub ProcEliminarProductor()
    With Form4
        eliminar = MsgBox("¿Desea eliminar el registro?", vbExclamation + vbYesNo, "Confirmación de eliminación")
        If eliminar = vbYes Then
            With BdRecordSet01
                'Eliminamos el registro
                .Delete
                .Requery
            End With
            ProcHabilitarBusqueda
            msg = MsgBox("Registro eliminado", vbOKOnly, "Terminado")
        End If
    End With
End Sub
Sub ProcEliminarValoresTabla()
    With Form5
        eliminar = MsgBox("¿Desea eliminar el registro?", vbExclamation + vbYesNo, "Confirmación de eliminación")
        If eliminar = vbYes Then
            With BdRecordSet01
                'Eliminamos el registro
                .Delete
                .Requery
            End With
            With .DataGrid1
                .SetFocus
            End With
            msg = MsgBox("Registro eliminado", vbOKOnly, "Terminado")
        End If
    End With
End Sub
Sub ProcFormResize()
    With Form1
        'Establecemos tamaño mìnimo del formulario
        If .Width < 5760 Then
            .Width = 5760
        End If
        If .Height < 5475 Then
            .Height = 5475
        End If
        'Posicionamos reloj
        With .Label1
            .Left = Form1.Width - 3345
        End With
        With .Label1
            .Top = Form1.Height - 1275
        End With
        'Redimensionamos el fondo
        With .Image1
            .Width = Form1.Width
        End With
        With .Image1
            .Height = Form1.Height
        End With
    End With
End Sub
Sub ProcHabilitarBusqueda()
    With Form4
        'Habilitamos el frame de búsqueda
        With .Frame1
            .Visible = False
        End With
        With .Frame2
            .Visible = False
        End With
        With .Frame3
            .Visible = True
        End With
        With .Text9
            .Text = ""
            .SetFocus
        End With
        With .Command1
            .Enabled = False
        End With
        With .Command2
            .Enabled = False
        End With
    End With
    BdRecordSet02.Requery
End Sub
Sub ProcHabilitarHato()
    With Form1
        'Habilitamos la selecciòn de un Hato, Dar de Alta Hatos y personal si ya se seleccionò el productor
        With .Label2
            If .Caption = "" Then
                With Form1
                    With .Label3
                        .Caption = ""
                    End With
                    VarHato = 0
                    With .SeleccionarHato
                        .Enabled = False
                    End With
                    With .Hato
                        .Enabled = False
                    End With
                    With .Personal
                        .Enabled = False
                    End With
                End With
            Else
                With Form1
                    With .SeleccionarHato
                        .Enabled = True
                    End With
                    With .Hato
                        .Enabled = True
                    End With
                    With .Personal
                        .Enabled = True
                    End With
                End With
            End If
        End With
    End With
End Sub
Sub ProcHatos()
    VarTipoConexionProductor = 2 'Hatos
    ProcConexionNuevoProductor
End Sub
Sub ProcInsertarProductor()
    With Form4
        With BdRecordSet01
            If VarTipoConexionProductor = 1 Then 'Productor
                insertar = MsgBox("¿Desea guardar el registro?", vbExclamation + vbYesNo, "Confirmación")
                'Insertamos el registro
                If insertar = vbYes Then
                    .AddNew
                    .Fields("Nombre") = Form4.Text1.Text
                    .Fields("Direccion") = Form4.Text2.Text
                    .Fields("Ciudad") = Form4.Text3.Text
                    .Fields("Creacion") = VarFecha
                    .Update
                    msg = MsgBox("Registro guardado correctamente", vbOKOnly, "Terminado")
                End If
            Else
                If VarTipoConexionProductor = 2 Then 'Hatos
                    insertar = MsgBox("¿Desea guardar el registro?", vbExclamation + vbYesNo, "Confirmación")
                    'Insertamos el registro
                    If insertar = vbYes Then
                        .AddNew
                        .Fields("Productor") = VarProductor
                        .Fields("Nombre") = Form4.Text1.Text
                        .Fields("Direccion") = Form4.Text2.Text
                        .Fields("Ciudad") = Form4.Text3.Text
                        .Fields("Creacion") = VarFecha
                        .Update
                        msg = MsgBox("Registro guardado correctamente", vbOKOnly, "Terminado")
                    End If
                Else
                    If VarTipoConexionProductor = 3 Then 'Personal
                        insertar = MsgBox("¿Desea guardar el registro?", vbExclamation + vbYesNo, "Confirmación")
                        'Insertamos el registro
                        If insertar = vbYes Then
                            .AddNew
                            .Fields("Productor") = VarProductor
                            .Fields("Nombre") = Form4.Text1.Text
                            .Fields("Direccion") = Form4.Text2.Text
                            .Fields("Ciudad") = Form4.Text3.Text
                            .Fields("Creacion") = VarFecha
                            .Update
                            msg = MsgBox("Registro guardado correctamente", vbOKOnly, "Terminado")
                        End If
                    End If
                End If
            End If
        End With
    End With
    ProcMostrarAgregarProductores
    With Form1
        'Deshabilitamos animales
        With .Animales
            .Enabled = False
        End With
        'Limpiamos etiquetas
        With .Label2
            .Caption = ""
        End With
        With .Label3
            .Caption = ""
        End With
        'Inicializamos variables
        VarProductor = 0
        VarHato = 0
        'Inabilitamos el seleccionar hatos
        With .SeleccionarHato
            .Enabled = False
        End With
    End With
End Sub
Sub ProcLimpiarFormulario()
    VarForm4State = 1
    With Form4
        'Inhabilitamos frames
        With .Frame1
            .Visible = False
        End With
        With .Frame2
            .Visible = False
        End With
        With .Frame3
            .Visible = True
        End With
        .Show
    End With
End Sub
Sub ProcListarTablas()
    With Form5
        With .Combo1
            'Agregamos valores al combobox
            .AddItem ("Categoría")
            .AddItem ("Causa de aborto")
            .AddItem ("Causa de no inseminar")
            .AddItem ("Causa de rechazo")
            .AddItem ("Diagnóstico de útero")
            .AddItem ("Enfermedad de ovario")
            .AddItem ("Enfermedad del útero")
            .AddItem ("Especificaciones de la muerte")
            .AddItem ("Especificaciones de venta")
            .AddItem ("Estado de la cría")
            .AddItem ("Medicamentos de cuartos mamarios")
            .AddItem ("Medicación Genital")
            .AddItem ("Medicamentos")
            .AddItem ("Razas")
            .AddItem ("Resultado análisis")
            .AddItem ("Sexo cría")
            .AddItem ("Tipo de análisis")
            .AddItem ("Tipo de enfermedad")
            .AddItem ("Tipo de parto")
            .AddItem ("Via de Aplicación")
        End With
    End With
End Sub
Sub ProcMostrarAgregarProductores()
    With Form4
        'Mostramos frame para insertar registro
        With .Frame1
            .Visible = True
        End With
        With .Frame2
            .Visible = False
        End With
        With .Frame3
            .Visible = False
        End With
        With .Text1
            .Text = ""
            .SetFocus
        End With
        With .Text2
            .Text = ""
        End With
        With .Text3
            .Text = ""
        End With
    End With
End Sub
Sub ProcPersonales()
    VarTipoConexionProductor = 3 'Personal
    ProcConexionNuevoProductor
End Sub
Sub ProcProductores()
    VarTipoConexionProductor = 1 'Productor
    ProcConexionNuevoProductor
End Sub
Sub ProcReloj()
    VarFecha = Date + Time
    With Form1
        With .Label1
            .Caption = VarFecha 'Reloj del Sistema
        End With
    End With
End Sub
Sub ProcResizeGrid()
    'Ajustamos el tamaño de las columnas de los diferentes grid
    If VarForm2State = 1 Then
        With Form2
            With .DataGrid1
                'Cambiamos el tamaño de las columnas del grid en el form2
                .Columns(0).Width = 500
                .Columns(1).Width = 5300
            End With
        End With
    End If
    If VarForm4State = 1 Then
        With Form4
            With .DataGrid3
                .Columns(0).Width = 500
                .Columns(1).Width = 5150
            End With
        End With
    End If
    If VarForm5State = 1 Then
        With Form5
            With .DataGrid1
                'Cambiamos el tamaño de las columnas del grid en el form2
                .Columns(0).Width = 700
                .Columns(1).Width = 5100
                .Columns(2).Width = 1400
            End With
        End With
    End If
End Sub
Sub ProcRespaldar()
    ProcSalirRespaldarRestaurar
    'Abrimos el cuadro de diálogo
    With Form1
        With .CommonDialog1
            .DialogTitle = "Guardar Respaldo"
            .Filter = "Respaldos|*.jahg"
            .InitDir = "C:\JAHG Software\Control Bovino\Respaldos"
            .ShowSave
            'Copiamos el archivo respaldo
            If .FileName = "" Then
            Else
                FileCopy "C:\JAHG Software\Control Bovino\BD\BD.mdb", .FileName
                msg = MsgBox("Respaldo terminado", vbOKOnly, "Terminado")
            End If
        End With
    End With
    ProcAbrirConexion
End Sub
Sub ProcRestaurar()
    ProcSalirRespaldarRestaurar
    'Abrimos el cuadro de diálogo
    With Form1
        With .CommonDialog1
            .DialogTitle = "Seleccionar Respaldo"
            .Filter = "Respaldos|*.jahg"
            .InitDir = "C:\JAHG Software\Control Bovino\Respaldos"
            .ShowOpen
            'Copiamos el archivo respaldo
            If .FileName = "" Then
            Else
                FileCopy .FileName, "C:\JAHG Software\Control Bovino\BD\BD.mdb"
                msg = MsgBox("Restauración terminada", vbOKOnly, "Terminado")
                main
            End If
        End With
    End With
End Sub
Sub ProcResultadoBusquedaProductores()
    With Form4
        'Mostramos los resultados de la búsqueda
        With .Frame1
            .Visible = False
        End With
        With .Frame2
            .Visible = True
        End With
        With .Frame3
            .Visible = False
        End With
        With .Text4
            .SetFocus
        End With
        With .Command1
            .Enabled = True
        End With
        With .Command2
            .Enabled = True
        End With
    End With
End Sub
Sub ProcSalir()
    VarForm6State = 0
    VarForm5State = 0
    VarForm4State = 0
    VarForm3State = 0
    VarForm2State = 0
    VarForm1State = 0
    ProcCerrarConexion
    'Cerramos todos los formularios
    Unload Form6
    Unload Form5
    Unload Form4
    Unload Form3
    Unload Form2
    Unload Form1
End Sub
Sub ProcSalirRespaldarRestaurar()
    VarForm6State = 0
    VarForm5State = 0
    VarForm4State = 0
    VarForm3State = 0
    VarForm2State = 0
    ProcCerrarConexion
    'Cerramos formularios para respaldar/restaurar
    Unload Form6
    Unload Form5
    Unload Form4
    Unload Form3
    Unload Form2
End Sub
Sub ProcSeleccionHato()
    With Form1
        'Deshabilitamos animales
        With .Animales
            .Enabled = False
        End With
        'Limpiamos etiqueta
        With .Label3
            .Caption = ""
        End With
    End With
    VarHato = 0
    VarForm2State = 1
    With Form2
        'Cambiamos el nombre del form2
        .Show
        .Caption = "Hatos de " + Form1.Label2.Caption
    End With
    'Asignamos valor a la variable
    VarTipo = "Hatos"
    'Àbrimos la conexiòn a la base de datos y se la asignamos al grid
    With BdRecordSet01
        If .State = 1 Then .Close
        .Open "select Id,Nombre from Hatos where productor = " & VarProductor, BdConexion, adOpenStatic, adLockOptimistic
        .Requery
    End With
    With Form2
        With .DataGrid1
            Set .DataSource = BdRecordSet01
        End With
    End With
    ProcResizeGrid
End Sub
Sub ProcSeleccionarImagenFondo()
    With Form1
        With .CommonDialog1
            'Abrimos el cuadro de diálogo
            .DialogTitle = "Seleccionar una Imagen de Fondo"
            .Filter = "Imágenes|*.jpg"
            .InitDir = "c:\%UserProfile%\Pictures"
            .ShowOpen
            'Asignamos la ruta a la variable
            VarFondo = .FileName
            If .FileName = "" Then
            Else
                ProcActualizarFondo
                ProcCargarFondo
            End If
        End With
    End With
End Sub
Sub ProcSeleccionProductor()
    With Form1
        'Deshabilitamos animales
        With .Animales
            .Enabled = False
        End With
        'Limpiamos etiquetas
        With .Label2
            .Caption = ""
        End With
        With .Label3
            .Caption = ""
        End With
        'Inicializamos variables
        VarProductor = 0
        VarHato = 0
        'Inabilitamos el seleccionar hatos
        With .SeleccionarHato
            .Enabled = False
        End With
    End With
    VarForm2State = 1
    With Form2
        'Cambiamos el nombre del form2
        .Show
        .Caption = "Productores"
        'Asignamos valor a la variable
        VarTipo = "Productores"
        'Àbrimos la conexiòn a la base de datos y se la asignamos al grid
        With BdRecordSet01
            If .State = 1 Then .Close
            .Open "select Id,Nombre from Productores", BdConexion, adOpenStatic, adLockOptimistic
            .Requery
        End With
        With .DataGrid1
            Set .DataSource = BdRecordSet01
        End With
    End With
    ProcResizeGrid
End Sub
