Attribute VB_Name = "Conexion"
Global Cn As New ADODB.Connection
Global RSIDNUM As New ADODB.Recordset
Global RsSomatometria As New ADODB.Recordset
Global RsLaboratorio As New ADODB.Recordset
Global RsDental As New ADODB.Recordset
Global RsNutricion As New ADODB.Recordset
Global RsSaludMujer As New ADODB.Recordset
Global RsOptometriaAudiometria As New ADODB.Recordset
Global RSTrabajador As New ADODB.Recordset
Global RsMujer As New ADODB.Recordset
Global RsMAudiometria As New ADODB.Recordset
Global RsTuberculosis As New ADODB.Recordset
Global RsCardio As New ADODB.Recordset
Global RsAllDate As New ADODB.Recordset
Global RsNombre As New ADODB.Recordset
Global RsCodigo As New ADODB.Recordset
Global RsFamiliar As New ADODB.Recordset
Global RsUsers As New ADODB.Recordset
Global RRsUsers As New ADODB.Recordset
Global RSSexo As New ADODB.Recordset
Global RsDocma As New ADODB.Recordset
Global RsDoccu As New ADODB.Recordset
Global RsMastografia As New ADODB.Recordset
Global RsAudiometria As New ADODB.Recordset
Global RSOptometria As New ADODB.Recordset
Global RSATuberculosis As New ADODB.Recordset
Global RSACardiologia As New ADODB.Recordset
Global RsGSomatometria As New ADODB.Recordset
Global RsGLaboratorio As New ADODB.Recordset
Global RsGDental As New ADODB.Recordset
Global RsGNutricion As New ADODB.Recordset
Global RsGAudiometria As New ADODB.Recordset
Global RSGOptometria As New ADODB.Recordset
Global RsGTuberculosis As New ADODB.Recordset
Global RSGCardiologia As New ADODB.Recordset
Global RsGMSomatometria As New ADODB.Recordset
Global RsGMLaboratorio As New ADODB.Recordset
Global RsGMDental As New ADODB.Recordset
Global RsGMNutricion As New ADODB.Recordset
Global RsGMAudiometria As New ADODB.Recordset
Global RSGMOptometria As New ADODB.Recordset
Global RSGMTuberculosis As New ADODB.Recordset
Global RSGMCardiologia As New ADODB.Recordset
Global ESO As New ADODB.Recordset
Global ELO As New ADODB.Recordset
Global EDO As New ADODB.Recordset
Global ENO As New ADODB.Recordset
Global EDMO As New ADODB.Recordset
Global EDCO As New ADODB.Recordset
Global EMGO As New ADODB.Recordset
Global EOO As New ADODB.Recordset
Global EAO As New ADODB.Recordset
Global ETO As New ADODB.Recordset
Global ECO As New ADODB.Recordset
Global ES3O As New ADODB.Recordset
Global EL3O As New ADODB.Recordset
Global ED3O As New ADODB.Recordset
Global EN3O As New ADODB.Recordset
Global EDM3O As New ADODB.Recordset
Global EDC3O As New ADODB.Recordset
Global EMG3O As New ADODB.Recordset
Global EO3O As New ADODB.Recordset
Global EA3O As New ADODB.Recordset
Global ET3O As New ADODB.Recordset
Global EC3O As New ADODB.Recordset
Global ES6O As New ADODB.Recordset
Global EL6O As New ADODB.Recordset
Global ED6O As New ADODB.Recordset
Global EN6O As New ADODB.Recordset
Global EDM6O As New ADODB.Recordset
Global EDC6O As New ADODB.Recordset
Global EMG6O As New ADODB.Recordset
Global EO6O As New ADODB.Recordset
Global EA6O As New ADODB.Recordset
Global ET6O As New ADODB.Recordset
Global EC6O As New ADODB.Recordset
Global TSOM As New ADODB.Recordset
Global TLAB As New ADODB.Recordset
Global TDEN As New ADODB.Recordset
Global TNUT As New ADODB.Recordset
Global TCMA As New ADODB.Recordset
Global TDOC As New ADODB.Recordset
Global TMAS As New ADODB.Recordset
Global TOPT As New ADODB.Recordset
Global TAUD As New ADODB.Recordset
Global TTUB As New ADODB.Recordset
Global TCAR As New ADODB.Recordset
Global ISOM As New ADODB.Recordset
Global ILAB As New ADODB.Recordset
Global IDEN As New ADODB.Recordset
Global INUT As New ADODB.Recordset
Global ICMA As New ADODB.Recordset
Global IDOC As New ADODB.Recordset
Global IMAS As New ADODB.Recordset
Global IOPT As New ADODB.Recordset
Global IAUD As New ADODB.Recordset
Global ITUB As New ADODB.Recordset
Global ICAR As New ADODB.Recordset
Global DRID_AST As String
Global DRNOMBRE As String
Global DRFECHA_NACIMIENTO As String
Global DRGENERO As String
Global DRPESO As String
Global DRTALLA As String
Global DRTA As String
Global DRVACUNA_TOXOIDE As String
Global DROTRAS_VACUNAS As String
Global DROBSERVACIONES_SOMATOMETRÍA As String
Global DRCOLESTEROL As String
Global DRTRIGLICERIDOS As String
Global DRGLUCOSA As String
Global DRPSA As String
Global DROBSERVACIONES_LABORATORIO As String
Global DRASISTENCIA_DENTAL As String
Global DROBSERVACIONES_DENTAL As String
Global DRASISTENCIA_NUTRICION As String
Global DRTIPO As String
Global DROBSERVACIONES_NUTRICION As String
Global DRDOCMA As String
Global DRDOCCU As String
Global DRMASTOGRAFIA As String
Global DROBSERVACIONES_SALUD_DE_LA_MUJER As String
Global DRAUDIOMETRIA As String
Global DROPTOMETRIA As String
Global DROBSERVACIONES_OPTOMETRIA As String
Global DRASISTENCIA_AUDIOMETRIA As String
Global DROBSERVACIONES_AUDIOMETRIA As String
Global DRASISTENCIA_TUBERCULOSIS As String
Global DROBSERVACIONES_TUBERCULOSIS As String
Global DRASISTENCIA_CARDIOLOGIA As String
Global DROBSERVACIONES_CARDIOLOGIA As String
Sub Main()
    With Cn
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=SQLPass;Persist Security Info=True;User ID=sa;Initial Catalog=SEMSAL;Data Source=SQLSERVER\SQLEXPRESS;"
    End With
    frmSplash.Show
End Sub
