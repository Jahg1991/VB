Attribute VB_Name = "Module1"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const IMAGE_BITMAP = 0
Public Const LR_COPYRETURNORG = &H4
Public Const CF_BITMAP = 2
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

Public vPrimeNombre As String
Public vSegundoNombre As String
Public vApellidoPaterno As String
Public vApellidoMaterno As String
Public vEstadoCivil As String
Public vNacionalidad As String
Public vSexo As String
Public vRfc As String
Public vCurp As String
Public vIfe As String
Public vImss As String
Public vCorreo As String
Public vTelefono As String
Public vTipoSangre As String
Public vEscolaridad As String
Public vEspecialidad As String
Public vCedulaProfesional As String
Public vTallaCamisa As String
Public vTallaPantalon As String
Public vTallaZapatos As String
Public vNombreFamiliar(0 To 18) As String
Public vParentescoFamiliar(0 To 18) As String
Public vDomicilioFamiliar(0 To 18) As String
Public vTelefonoFamiliar(0 To 18) As String
Public vLesionDetalle As String
Public vInternadoDetalle As String
Public vCronicas As String
Public vMedicamento As String
Public vAlergias As String
Public vCigarrillos As String
Public vIdiomas As String
Public vProgramas As String
Public vDetalleConducir As String
Public vAreaLaboro As String
Public vTiempoLaboro As String
Public vMotivoRenuncia As String
Public vEmpresaAnterior(0 To 1) As String
Public vDomicilioEmpresaAnterior(0 To 1) As String
Public vTiempoEmpresaAnterior(0 To 1) As String
Public vEncargadoEmpresaAnterior(0 To 1) As String
Public vActividadesEmpresaAnterior(0 To 1) As String
Public vSueldoEmpresaAnterior(0 To 1) As String
Public vRenunciaEmpresaAnterior(0 To 1) As String
Public vPuestosolicita As String
Public vAptoPuesto As String
Public vHabilidadesPuesto As String
Public vSalarioEsperado As String
Public vTipoCompaneros As String
Public vConocioEmpresa As String
Public vHobbie As String
Public vPositivo As String
Public vNegativo As String
Public vCalle As String
Public vExterior As String
Public vInterior As String
Public vColonia As String
Public vCiudad As String
Public vCodigoPostal As String
Public vEstado As String
Public vPais As String
Public vFoto As String
    
Public vFechaNacimiento As Date
Public vnacimientoFamiliar(0 To 18) As Date
    
Public vLesion As Boolean
Public vInternado As Boolean
Public vFuma As Boolean
Public vAlcohol As Boolean
Public vLeer As Boolean
Public vEscribir As Boolean
Public vManejar As Boolean
Public vLaboroAnteriormente As Boolean
Public vFamiliaresLaborando As Boolean
Public vPoseeVehiculo As Boolean
Public vPoseeMascotas As Boolean
Public vOtrosIngresos As Boolean
Public vCasaPropia As Boolean
Public vParejaTrabaja As Boolean
Public vDeudas As Boolean
Public vDeporte As Boolean
Public vPrincipios As Boolean

Sub Limpiar()

    Dim i As Integer
    
    With Form1
    
        .SSTab1.Tab = 0
    
        .DTPicker1.Value = Date
        
        i = 0
        Do While i <= 18
            .DTPicker2(i).Value = Date
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 12
            .Text1(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 56
            .Text2(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 5
            .Text3(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 2
            .Text4(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 16
            .Text5(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 4
            .Text6(i).Text = ""
            i = i + 1
        Loop
        
        .Text7.Text = ""
        
        i = 0
        Do While i <= 2
            .Text8(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 6
            .Text9(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 5
            .Combo1(i).Text = ""
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 18
            .Combo2(i).Text = ""
            i = i + 1
        Loop
        
        .Combo3.Text = ""
        
        i = 0
        Do While i <= 3
            .Check1(i).Value = 0
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 2
            .Check2(i).Value = 0
            i = i + 1
        Loop
        
        .Check3.Value = 0
        
        i = 0
        Do While i <= 2
            .Check4(i).Value = 0
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 3
            .Check5(i).Value = 0
            i = i + 1
        Loop
        
        i = 0
        Do While i <= 1
            .Check6(i).Value = 0
            i = i + 1
        Loop
    
    End With
    
    vPrimeNombre = ""
    vSegundoNombre = ""
    vApellidoPaterno = ""
    vApellidoMaterno = ""
    vEstadoCivil = ""
    vNacionalidad = ""
    vFechaNacimiento = Date
    vSexo = ""
    vRfc = ""
    vCurp = ""
    vIfe = ""
    vImss = ""
    vCorreo = ""
    vTelefono = ""
    vTipoSangre = ""
    vEscolaridad = ""
    vEspecialidad = ""
    vCedulaProfesional = ""
    vTallaPantalon = ""
    vTallaCamisa = ""
    vTallaZapatos = ""
    vFoto = ""
        
    i = 0
    Do While i <= 18
        vNombreFamiliar(i) = ""
        vParentescoFamiliar(i) = ""
        vDomicilioFamiliar(i) = ""
        vTelefonoFamiliar(i) = ""
        vnacimientoFamiliar(i) = Date
        i = i + 1
    Loop
        
    vLesion = 0
    vLesionDetalle = ""
    vInternado = 0
    vInternadoDetalle = ""
    vCronicas = ""
    vMedicamento = ""
    vAlergias = ""
    vFuma = 0
    vCigarrillos = ""
    vAlcohol = 0
        
    vIdiomas = ""
    vProgramas = ""
    vLeer = 0
    vEscribir = 0
    vManejar = 0
    vDetalleConducir = ""
        
        
    vLaboroAnteriormente = 0
    vAreaLaboro = ""
    vTiempoLaboro = ""
    vMotivoRenuncia = ""
        
    i = 0
    Do While i <= 1
        vEmpresaAnterior(i) = ""
        vDomicilioEmpresaAnterior(i) = ""
        vTiempoEmpresaAnterior(i) = ""
        vEncargadoEmpresaAnterior(i) = ""
        vActividadesEmpresaAnterior(i) = ""
        vSueldoEmpresaAnterior(i) = ""
        vRenunciaEmpresaAnterior(i) = ""
        i = i + 1
    Loop

    vPuestosolicita = ""
    vAptoPuesto = ""
    vHabilidadesPuesto = ""
    vSalarioEsperado = ""
    vTipoCompaneros = ""
        
    vConocioEmpresa = ""
    vFamiliaresLaborando = 0
    vPoseeVehiculo = 0
    vPoseeMascotas = 0
        
    vOtrosIngresos = 0
    vCasaPropia = 0
    vParejaTrabaja = 0
    vDeudas = 0
        
    vHobbie = ""
    vDeporte = 0
    vPrincipios = 0
    vPositivo = ""
    vNegativo = ""
    
    vCalle = ""
    vExterior = ""
    vInterior = ""
    vColonia = ""
    vCiudad = ""
    vCodigoPostal = ""
    vEstado = ""
    vPais = ""
    
    With Form1
        .Text1(9).Enabled = False
        .Text1(10).Enabled = False
        .Text3(0).Enabled = False
        .Text3(1).Enabled = False
        .Text3(5).Enabled = False
        .Text4(2).Enabled = False
        .Text5(0).Enabled = False
        .Text5(1).Enabled = False
        .Text5(2).Enabled = False
        .Check5(2).Enabled = False
    End With
    
    Form1.Image1.Picture = LoadPicture(vFoto)

End Sub

Sub CargarCombos()

    Dim i As Integer

    With Form1
    
        With .Combo1(0)
            .AddItem "CASADO"
            .AddItem "DIVORCIADO"
            .AddItem "SOLTERO"
            .AddItem "UNION LIBRE"
            .AddItem "VIUDO"
        End With
        
        With .Combo1(1)
            .AddItem "AMERICANA"
            .AddItem "CANADIENSE"
            .AddItem "ESPAÑOLA"
            .AddItem "MEXICANA"
            .AddItem "VENOZALANA"
        End With
        
        With .Combo1(2)
            .AddItem "FEMENINO"
            .AddItem "INDEFINIDO"
            .AddItem "MASCULINO"
        End With
        
        With .Combo1(3)
            .AddItem ""
            .AddItem "A+"
            .AddItem "A-"
            .AddItem "AB+"
            .AddItem "AB-"
            .AddItem "B+"
            .AddItem "B-"
            .AddItem "O+"
            .AddItem "O-"
        End With
        
        With .Combo1(4)
            .AddItem "DOCTORADO"
            .AddItem "MAESTRIA"
            .AddItem "NO ESTUDIOS"
            .AddItem "PREPARATORIA"
            .AddItem "PRIMARIA"
            .AddItem "PROFESIONAL"
            .AddItem "SECUNDARIA"
            .AddItem "TECNICA"
        End With
        
        With .Combo1(5)
            .AddItem "Chica S"
            .AddItem "Doble Extra Grande XXL"
            .AddItem "Extra Chica XS"
            .AddItem "Extra Grande XL"
            .AddItem "Grande L"
            .AddItem "Mediana M"
            .AddItem "Triple Extra Grande XXXL"
        End With
        
        i = 0
        Do While i <= 18
            With .Combo2(i)
                .AddItem "Abuelo"
                .AddItem "Amigo"
                .AddItem "Esposa"
                .AddItem "Esposo"
                .AddItem "Hermana"
                .AddItem "Hermano"
                .AddItem "Hija"
                .AddItem "Hijo"
                .AddItem "Madre"
                .AddItem "Nieto"
                .AddItem "Padre"
                .AddItem "Pareja"
                .AddItem "Primo"
                .AddItem "Sobrino"
                .AddItem "Suegro"
                .AddItem "Tío"
            End With
            i = i + 1
        Loop
        
        With .Combo3
            .AddItem "Aguascalientes"
            .AddItem "Baja California"
            .AddItem "Baja California Sur"
            .AddItem "Campeche"
            .AddItem "Chiapas"
            .AddItem "Chihuahua"
            .AddItem "Ciudad de México"
            .AddItem "Coahuila de Zaragoza"
            .AddItem "Colima"
            .AddItem "Durango"
            .AddItem "Guanajuato"
            .AddItem "Guerrero"
            .AddItem "Hidalgo"
            .AddItem "Jalisco"
            .AddItem "México"
            .AddItem "Michoacán de Ocampo"
            .AddItem "Morelos"
            .AddItem "Nayarit"
            .AddItem "Nuevo León"
            .AddItem "Oaxaca"
            .AddItem "Puebla"
            .AddItem "Querétaro de Arteaga"
            .AddItem "Quintana Roo"
            .AddItem "San Luis Potosí"
            .AddItem "Sinaloa"
            .AddItem "Sonora"
            .AddItem "Tabasco"
            .AddItem "Tamaulipas"
            .AddItem "Tlaxcala"
            .AddItem "Veracruz de Ignacio de la Llave"
            .AddItem "Yucatán"
            .AddItem "Zacatecas"
        End With
        
    End With

End Sub

Sub CrearTxt()

On Error GoTo ErrorGeneral

    Dim i As Integer
    Dim t As Integer
    Dim u As Integer
    Dim v As Integer
    Dim w As Integer
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    
    Dim NombreArchivo, RutaArchivo As String
    Dim Obj As FileSystemObject
    Dim tx As Scripting.TextStream
    
    With Form1
        
        vPrimeNombre = .Text1(0)
        vSegundoNombre = .Text1(1)
        vApellidoPaterno = .Text1(2)
        vApellidoMaterno = .Text1(3)
        vEstadoCivil = .Combo1(0)
        vNacionalidad = .Combo1(1)
        vFechaNacimiento = .DTPicker1
        vSexo = .Combo1(2)
        vRfc = .Text1(4)
        vCurp = .Text1(5)
        vIfe = .Text1(6)
        vImss = .Text1(7)
        vCorreo = .Text1(8)
        vTelefono = .Text1(13)
        vTipoSangre = .Combo1(3)
        vEscolaridad = .Combo1(4)
        vEspecialidad = .Text1(9)
        vCedulaProfesional = .Text1(10)
        vTallaPantalon = .Text1(11)
        vTallaCamisa = .Combo1(5)
        vTallaZapatos = .Text1(12)
        
        i = 0
        x = 0
        y = 1
        z = 2
        Do While i <= 18
            vNombreFamiliar(i) = .Text2(x)
            vParentescoFamiliar(i) = .Combo2(i)
            vDomicilioFamiliar(i) = .Text2(y)
            vTelefonoFamiliar(i) = .Text2(z)
            vnacimientoFamiliar(i) = .DTPicker2(i)
            i = i + 1
            x = x + 3
            y = y + 3
            z = z + 3
        Loop
        
        vLesion = .Check1(0)
        vLesionDetalle = .Text3(0)
        vInternado = .Check1(1)
        vInternadoDetalle = .Text3(1)
        vCronicas = .Text3(2)
        vMedicamento = .Text3(3)
        vAlergias = .Text3(4)
        vFuma = .Check1(2)
        vCigarrillos = .Text3(5)
        vAlcohol = .Check1(3)
        
        vIdiomas = .Text4(0)
        vProgramas = .Text4(1)
        vLeer = .Check2(0)
        vEscribir = .Check2(1)
        vManejar = .Check2(2)
        vDetalleConducir = .Text4(2)
        
        vLaboroAnteriormente = .Check3
        vAreaLaboro = .Text5(0)
        vTiempoLaboro = .Text5(1)
        vMotivoRenuncia = .Text5(2)
        
        i = 0
        t = 3
        u = 4
        v = 5
        w = 6
        x = 7
        y = 8
        z = 9
        Do While i <= 1
            vEmpresaAnterior(i) = .Text5(t)
            vDomicilioEmpresaAnterior(i) = .Text5(u)
            vTiempoEmpresaAnterior(i) = .Text5(v)
            vEncargadoEmpresaAnterior(i) = .Text5(w)
            vActividadesEmpresaAnterior(i) = .Text5(x)
            vSueldoEmpresaAnterior(i) = .Text5(y)
            vRenunciaEmpresaAnterior(i) = .Text5(z)
            i = i + 1
            t = t + 7
            u = u + 7
            v = v + 7
            w = w + 7
            x = x + 7
            y = y + 7
            z = z + 7
        Loop

        vPuestosolicita = .Text6(0)
        vAptoPuesto = .Text6(1)
        vHabilidadesPuesto = .Text6(2)
        vSalarioEsperado = .Text6(3)
        vTipoCompaneros = .Text6(4)
        
        vConocioEmpresa = .Text7
        vFamiliaresLaborando = .Check4(0)
        vPoseeVehiculo = .Check4(1)
        vPoseeMascotas = .Check4(2)
        
        vOtrosIngresos = .Check5(0)
        vCasaPropia = .Check5(1)
        vParejaTrabaja = .Check5(2)
        vDeudas = .Check5(3)
        
        vHobbie = .Text8(0)
        vDeporte = .Check6(0)
        vPrincipios = .Check6(1)
        vPositivo = .Text8(1)
        vNegativo = .Text8(2)
        
        vCalle = .Text9(0)
        vExterior = .Text9(1)
        vInterior = .Text9(2)
        vColonia = .Text9(3)
        vCiudad = .Text9(4)
        vCodigoPostal = .Text9(5)
        vEstado = .Combo3
        vPais = .Text9(6)
        
    End With
    
    If vPrimeNombre = "" Then GoTo NombreVacio Else
        If vApellidoPaterno = "" Then GoTo ApellidoPaternoVacio Else
            If vApellidoMaterno = "" Then GoTo ApellidoMaternoVacio Else
                If vCalle = "" Then GoTo CalleVacio Else
                    If vCiudad = "" Then GoTo CiudadVacio Else
                        If vEstado = "" Then GoTo EstadoVacio Else
                            If vPais = "" Then GoTo PaisVacio Else
                                If vFoto = "" Then GoTo FotoVacia Else
    
                                    If vSegundoNombre = "" Then
                                        NombreArchivo = vPrimeNombre & " " & vApellidoPaterno & " " & vApellidoMaterno
                                    Else
                                        NombreArchivo = vPrimeNombre & " " & vSegundoNombre & " " & vApellidoPaterno & " " & vApellidoMaterno
                                    End If
                                    
                                    NombreArchivo = Replace(NombreArchivo, "Ñ", "N")
                                    
                                    NombreArchivo = Replace(NombreArchivo, "ñ", "n")
                                    
                                    On Error Resume Next
                                    'filecopy "C:\JAHG Software\Aspirantes\image.tif", "D:\" & NombreArchivo & ".tif"
                                    FileCopy "C:\JAHG Software\Aspirantes\image.tif", "\\10.2.2.248\Interfas\Cord\Aspirantes\bmp\" & NombreArchivo & ".tif"
                                    
                                    'RutaArchivo = "D:\" & NombreArchivo & ".csv"
                                    RutaArchivo = "\\10.2.2.248\Interfas\Cord\Aspirantes\txt\" & NombreArchivo & ".csv"
                                        
                                    Set Obj = New FileSystemObject
                                    Set tx = Obj.CreateTextFile(RutaArchivo)
                                    
                                    tx.Write "1"
                                    tx.Write "|"
                                    tx.Write vPrimeNombre
                                    tx.Write "|"
                                    tx.Write vSegundoNombre
                                    tx.Write "|"
                                    tx.Write vApellidoPaterno
                                    tx.Write "|"
                                    tx.Write vApellidoMaterno
                                    tx.Write "|"
                                    tx.Write vEstadoCivil
                                    tx.Write "|"
                                    tx.Write vNacionalidad
                                    tx.Write "|"
                                    tx.Write vFechaNacimiento
                                    tx.Write "|"
                                    tx.Write vSexo
                                    tx.Write "|"
                                    tx.Write vRfc
                                    tx.Write "|"
                                    tx.Write vCurp
                                    tx.Write "|"
                                    tx.Write vIfe
                                    tx.Write "|"
                                    tx.Write vImss
                                    tx.Write "|"
                                    tx.Write vCorreo
                                    tx.Write "|"
                                    tx.Write vTipoSangre
                                    tx.Write "|"
                                    tx.Write vEscolaridad
                                    tx.Write "|"
                                    tx.Write vEspecialidad
                                    tx.Write "|"
                                    tx.Write vCedulaProfesional
                                    tx.Write "|"
                                    tx.Write vTallaPantalon
                                    tx.Write "|"
                                    tx.Write vTallaCamisa
                                    tx.Write "|"
                                    tx.Write vTallaZapatos
                                        
                                    tx.Write "|"
                                    tx.Write vLesion
                                    tx.Write "|"
                                    tx.Write vLesionDetalle
                                    tx.Write "|"
                                    tx.Write vInternado
                                    tx.Write "|"
                                    tx.Write vInternadoDetalle
                                    tx.Write "|"
                                    tx.Write vCronicas
                                    tx.Write "|"
                                    tx.Write vMedicamento
                                    tx.Write "|"
                                    tx.Write vAlergias
                                    tx.Write "|"
                                    tx.Write vFuma
                                    tx.Write "|"
                                    tx.Write vCigarrillos
                                    tx.Write "|"
                                    tx.Write vAlcohol
                                        
                                    tx.Write "|"
                                    tx.Write vIdiomas
                                    tx.Write "|"
                                    tx.Write vProgramas
                                    tx.Write "|"
                                    tx.Write vLeer
                                    tx.Write "|"
                                    tx.Write vEscribir
                                    tx.Write "|"
                                    tx.Write vManejar
                                        
                                    tx.Write "|"
                                    tx.Write vLaboroAnteriormente
                                    tx.Write "|"
                                    tx.Write vAreaLaboro
                                    tx.Write "|"
                                    tx.Write vTiempoLaboro
                                    tx.Write "|"
                                    tx.Write vMotivoRenuncia
                                        
                                    i = 0
                                    Do While i <= 1
                                        tx.Write "|"
                                        tx.Write vEmpresaAnterior(i)
                                        tx.Write "|"
                                        tx.Write vDomicilioEmpresaAnterior(i)
                                        tx.Write "|"
                                        tx.Write vTiempoEmpresaAnterior(i)
                                        tx.Write "|"
                                        tx.Write vEncargadoEmpresaAnterior(i)
                                        tx.Write "|"
                                        tx.Write vActividadesEmpresaAnterior(i)
                                        tx.Write "|"
                                        tx.Write vSueldoEmpresaAnterior(i)
                                        tx.Write "|"
                                        tx.Write vRenunciaEmpresaAnterior(i)
                                        i = i + 1
                                    Loop
                                
                                    tx.Write "|"
                                    tx.Write vPuestosolicita
                                    tx.Write "|"
                                    tx.Write vAptoPuesto
                                    tx.Write "|"
                                    tx.Write vHabilidadesPuesto
                                    tx.Write "|"
                                    tx.Write vSalarioEsperado
                                    tx.Write "|"
                                    tx.Write vTipoCompaneros
                                        
                                    tx.Write "|"
                                    tx.Write vConocioEmpresa
                                    tx.Write "|"
                                    tx.Write vFamiliaresLaborando
                                    tx.Write "|"
                                    tx.Write vPoseeVehiculo
                                    tx.Write "|"
                                    tx.Write vPoseeMascotas
                                        
                                    tx.Write "|"
                                    tx.Write vOtrosIngresos
                                    tx.Write "|"
                                    tx.Write vCasaPropia
                                    tx.Write "|"
                                    tx.Write vParejaTrabaja
                                    tx.Write "|"
                                    tx.Write vDeudas
                                        
                                    tx.Write "|"
                                    tx.Write vHobbie
                                    tx.Write "|"
                                    tx.Write vDeporte
                                    tx.Write "|"
                                    tx.Write vPrincipios
                                    tx.Write "|"
                                    tx.Write vPositivo
                                    tx.Write "|"
                                    tx.Write vNegativo
                                    tx.Write "|"
                                    tx.Write vTelefono
                                    
                                    tx.Write "|"
                                    tx.Write vCalle
                                    tx.Write "|"
                                    tx.Write vExterior
                                    tx.Write "|"
                                    tx.Write vInterior
                                    tx.Write "|"
                                    tx.Write vColonia
                                    tx.Write "|"
                                    tx.Write vCiudad
                                    tx.Write "|"
                                    tx.Write vCodigoPostal
                                    tx.Write "|"
                                    tx.Write vEstado
                                    tx.Write "|"
                                    tx.Write vPais
                                    tx.Write "|"
                                    tx.Write vDetalleConducir
                                    tx.Write "|"
                                    
                                    tx.WriteLine
                                        
                                    i = 0
                                    Do While i <= 18
                                        If vNombreFamiliar(i) <> "" And vParentescoFamiliar(i) <> "" Then
                                            tx.Write "2"
                                            tx.Write "|"
                                            tx.Write vNombreFamiliar(i)
                                            tx.Write "|"
                                            tx.Write vParentescoFamiliar(i)
                                            tx.Write "|"
                                            tx.Write vDomicilioFamiliar(i)
                                            tx.Write "|"
                                            tx.Write vTelefonoFamiliar(i)
                                            tx.Write "|"
                                            tx.Write vnacimientoFamiliar(i)
                                            tx.Write "|"
                                            tx.WriteLine
                                        End If
                                        i = i + 1
                                    Loop
                                    
                                    tx.Close
                                    Set Obj = Nothing
                                    FileCopy "\\10.2.2.248\Interfas\Cord\Aspirantes\txt\" & NombreArchivo & ".csv", "\\10.2.2.248\Interfas\Cord\Aspirantes\bck\" & NombreArchivo & ".csv"
                                    MsgBox "Archivo creado: " & RutaArchivo, vbOKOnly, "Terminado"
                                    Limpiar
                                    Exit Sub
    
ErrorGeneral:
    MsgBox "Comuníquese con el Departamento Sistemas Oracle", vbCritical, "Error"
    Exit Sub
    
NombreVacio:
    MsgBox "El primer nombre es obligatorio", vbOKOnly, "Error"
    Exit Sub

ApellidoPaternoVacio:
    MsgBox "El apellido paterno es obligatorio", vbOKOnly, "Error"
    Exit Sub

ApellidoMaternoVacio:
    MsgBox "El apellido materno es obligatorio", vbOKOnly, "Error"
    Exit Sub
    
CalleVacio:
    MsgBox "La calle en la dirección es obligatoria", vbOKOnly, "Error"
    Exit Sub

CiudadVacio:
    MsgBox "La ciudad en la dirección es obligatoria", vbOKOnly, "Error"
    Exit Sub

EstadoVacio:
    MsgBox "El estado en la dirección es obligatorio", vbOKOnly, "Error"
    Exit Sub

PaisVacio:
    MsgBox "El país en la dirección es obligatorio", vbOKOnly, "Error"
    Exit Sub
    
FotoVacia:
    MsgBox "La fotografía es obligatoria", vbOKOnly, "Error"
    Exit Sub

End Sub

Sub getSnapshot()

Dim RetVal, base, i, picName, picPath

    picName = "image.tif"

    ' Make sure the current directory is set to the one
    ' where the Excel file is saved
    base = "C:\JAHG Software\Aspirantes\"
    picPath = base & picName
    
    ' Capture new image
    On Error GoTo FileMissing
    RetVal = Shell(base & "CommandCam.exe /filename """ & picPath & """", vbHide)
    
    dsm "Espere 5 segundos mientras se procesa la fotografía"
    
    Calcular
    
    ' Load new image into image object on spreadsheet
    'dsm "Attempting to load the picture on a new tab "
    'NewPictureSheet picPath, "Picture " & picNumber
    
    dsm "Tomar foto"
    Form1.Image1.Picture = LoadPicture(picPath)
    vFoto = picPath
    
    Dim hNew2 As Long

    hNew2 = CopyImage(Form1.Image1.Picture, IMAGE_BITMAP, Val(100), Val(100), LR_COPYRETURNORG)
    OpenClipboard Form1.hwnd
    EmptyClipboard
    SetClipboardData CF_BITMAP, hNew2
    CloseClipboard

    Form1.Image1.Picture = Clipboard.GetData(2)
    SavePicture Form1.Image1.Picture, picPath
    
    Exit Sub
    
FileMissing:
    MsgBox "No se encuentra el archivo CommandCam.exe", vbOKOnly, "Error"
    Exit Sub

End Sub

Sub dsm(msg As String)
    Form1.Command2.Caption = msg
End Sub

Sub Calcular()
    Call Sleep(5000)
End Sub
