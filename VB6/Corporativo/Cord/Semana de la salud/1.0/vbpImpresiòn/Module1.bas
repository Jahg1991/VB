Attribute VB_Name = "Module1"
Global Cn As New ADODB.Connection
Global Rs As New ADODB.Recordset

Global VarNombre As String
Global VarFecha_nacimiento As String
Global VarGenero As String
Global VarPeso As String
Global VarTalla As String
Global VarTension_arterial As String
Global VarVacuna_toxoide As String
Global VarOtras_vacunas As String
Global VarObservaciones_somatometria As String
Global VarColesterol As String
Global VarTrigliceridos As String
Global VarGlucosa As String
Global VarObservaciones_laboratorio As String
Global VarLavado_oidos As String
Global VarPrueba_audicion As String
Global VarObservaciones_audiometria As String
Global VarCardiologia As String
Global VarLimpieza_dental As String
Global VarRevision_dental As String
Global VarObservaciones_dental As String
Global VarDoccu As String
Global VarDocm As String
Global Varmastografia As String
Global VarConsulta_nutricion As String
Global VarPlatica_nutricion As String
Global VarObservaciones_nutricion As String
Global VarObservaciones_optometria As String
Global VarObservaciones_tuberculosis As String

Sub main()
    With Cn
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=Santateresa1;Persist Security Info=True;User ID=ss16;Initial Catalog=ss16;Data Source=SQLSERVER\SQLEXPRESS;"
    End With
    With Rs
        If .State = 1 Then .Close
            .Open "select * from Reporte", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Form1.Show
End Sub

