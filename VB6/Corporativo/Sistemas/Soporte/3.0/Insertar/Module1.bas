Attribute VB_Name = "Module1"
Global CN As New ADODB.Connection

Global RSAGENDA As New ADODB.Recordset
Global RSAGENDA2 As New ADODB.Recordset
Global RSCOMUNICADOS As New ADODB.Recordset
Global RSMISION As New ADODB.Recordset
Global RSOBJETIVO As New ADODB.Recordset
Global RSREGISTRO As New ADODB.Recordset
Global RSRESENA As New ADODB.Recordset
Global RSSLASH As New ADODB.Recordset
Global RSSLASHID As New ADODB.Recordset
Global RSVISION As New ADODB.Recordset
Global RSID As New ADODB.Recordset

Global VARUSUARIO As String

' Declaraci�n del Api GetUserName
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long) As Long
  
' Retorna un String con el nombre de usuario actual de windows
' ************************************************************
Private Function get_Usuario() As String
      
    Dim Nombre As String, ret As Long
      
    ' Buffer
    Nombre = Space$(250)
      
    ' Tama�o
    ret = Len(Nombre)
      
    If GetUserName(Nombre, ret) = 0 Then
        get_Usuario = vbNullString
    Else
        ' Extrae solo los caracteres
        get_Usuario = Left$(Nombre, ret - 1)
    End If
      
End Function

Sub MAIN()
    
    'ABRIMOS LA CONEXION
    CN.CursorLocation = adUseClient
    CN.Open "Provider=SQLOLEDB.1;Password=Jahg1991;Persist Security Info=True;User ID=sa;Initial Catalog=SOPORTE;Data Source=EQUIPO05\JAHG;"
    
    'ABRIMOS LOS RECORSETS
    With RSAGENDA
        If .State = 1 Then .Close
            .Open "SELECT * FROM AGENDA", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSAGENDA2
        If .State = 1 Then .Close
            .Open "SELECT NOMBRE FROM AGENDA ORDER BY 1", CN, adOpenRead, adLockReadOnly
            .Requery
    End With
    
    With RSCOMUNICADOS
        If .State = 1 Then .Close
            .Open "SELECT * FROM COMUNICADOS ORDER BY 1 DESC", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSID
        If .State = 1 Then .Close
            .Open "SELECT ID FROM AGENDA ORDER BY 1 DESC", CN, adOpenRead, adLockReadOnly
            .Requery
    End With
    
    With RSMISION
        If .State = 1 Then .Close
            .Open "SELECT * FROM MISION", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSOBJETIVO
        If .State = 1 Then .Close
            .Open "SELECT * FROM OBJETIVO", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSREGISTRO
        If .State = 1 Then .Close
            .Open "SELECT * FROM REGISTRO", CN, adOpenStatic, adLockOptimistic
            .Requery
            'INSERTAMOS EL REGISTRO DE VISITA
            .AddNew
                .Fields("NOMBRE") = get_Usuario
                .Fields("FECHA") = Date + Time
            .Update
            .Requery
            .Close
            
            VARUSUARIO = get_Usuario
    End With
    
    With RSRESENA
        If .State = 1 Then .Close
            .Open "SELECT * FROM RESENA", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSSLASH
        If .State = 1 Then .Close
            .Open "SELECT * FROM SLASH", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    With RSSLASHID
        If .State = 1 Then .Close
            .Open "SELECT ID FROM SLASH ORDER BY 1 DESC", CN, adOpenRead, adLockReadOnly
            .Requery
    End With
    
    With RSVISION
        If .State = 1 Then .Close
            .Open "SELECT * FROM VISION", CN, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    Form1.Caption = "Bienvenido " + get_Usuario
    Form1.Show
    
End Sub
