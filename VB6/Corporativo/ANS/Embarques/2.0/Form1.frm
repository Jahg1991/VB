VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JAHG Software -  Embarques EBS"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3915
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   1440
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'Declaración de variables
'*********************************************************************************

Public cnn As New ADODB.Connection 'Conexión a la base de datos

Public Rst As New ADODB.Recordset  'Juego de registros datos de embarque
Public Rst2 As New ADODB.Recordset  'Juego de registros datos de embarque

Public C1 As String
Public C2 As String

Public R As Boolean

Public varcount As Integer  'Contador carpeta \\10.1.2.252\Embarques\pesadas\rev
Public varcount2 As Integer 'Contador carpeta \\10.1.2.252\Embarques\pesadas\in
Public varCuenta As Integer 'Identificar id archivo

Public Datos_Csv As String

Public ObjCarpeta As Object
Public Carpeta As Object
Public ObjCarpeta2 As Object
Public Carpeta2 As Object
    
Option Explicit
  
'Función que exporta el recordset a un Archivo de texto csv separado por comas
'*********************************************************************************

Public Function Recordset_a_CSV(rs As ADODB.Recordset, Path_Csv As String) As Boolean

    On Error GoTo errFunction
      
    ' Devuelve los datos separados por comas y con un salto de carro
    Datos_Csv = rs.GetString(adClipString, -1, "|", vbCrLf, "")
    Datos_Csv = Replace(Datos_Csv, vbCrLf, " ")
      
    ' Abre y Crea un archivo de texto para escribir los datos
    Open Path_Csv For Output As #1
    
    ' escribe los datos
    Print #1, Datos_Csv
    
    'cierra
    Close
    
    ' Ok
    Recordset_a_CSV = True
  
Exit Function

    'Error
errFunction:
  
    MsgBox Err.Description, vbCritical
  
End Function
      
Private Sub Form_Load()
          
        ' Nueva conexión Ado
        cnn.CursorLocation = adUseClient
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Basculas Revuelta\Revuelta 100 XXI 3.0\rev_2018.mdb;Persist Security Info=False"
        
        C1 = "SELECT " _
                & "boleto_ent.clave_e AS clave_e,boleto_ent.clave_c AS clave_cliente,cliente.nombre_c AS nombre_cliente,boleto_ent.clave_o AS clave_operador,operador.nombre_o AS nombre_operador," _
                & "boleto_ent.clave_p AS clave_producto,producto.nombre_p AS nombre_producto,boleto_ent.completo AS completo,boleto_ent.fecha_e AS fecha_entrada,boleto_ent.hora_e AS hora_entrada,boleto_ent.placas AS placas," _
                & "boleto_ent.conductor AS conductor,boleto_ent.peso_e AS peso_entrada,boleto_ent.observa_e AS observaciones_entrada,boleto_ent.unidad_e AS unidad_entrada," _
                & "boleto_ent.bascula_e AS bascula_entrada,boleto_ent.t_entrada AS t_entrada,boleto_ent.clave_u AS clave_u,boleto_ent.tipo_pesada AS tipo_pesada,boleto_ent.adicional1 AS adicional," _
                & "boleto_ent.precio_producto AS precio_producto,boleto_ent.tipo_precio AS tipo_precio,boleto_sal.fecha_s AS fecha_salida,boleto_sal.hora_s AS hora_salida," _
                & "boleto_sal.peso_s AS peso_salida,boleto_sal.peso_n AS peso_neto,boleto_sal.turno_s AS turno_salida,boleto_sal.bascula_s AS bascula_salida,boleto_sal.s_manual AS salida_manual," _
                & "boleto_sal.precio_total AS precio_total,boleto_sal.nombre_os AS nombre_os,boleto_sal.observa_s AS observa_salida " _
             & "From producto,operador ,cliente ,boleto_ent ,boleto_sal " _
             & "Where " _
                & "boleto_ent.clave_e = boleto_sal.clave_e " _
                & "AND boleto_ent.clave_c = cliente.clave_c " _
                & "AND boleto_ent.clave_o = operador.clave_o " _
                & "AND boleto_ent.clave_p = producto.clave_p " _
                & "AND cliente.nombre_c = 'AGROPECUARIA NUEVO SIGLO S.A.de C.V.' " _
                & "ORDER BY boleto_ent.clave_e DESC "
                
        C2 = "SELECT top 1 " _
                & "boleto_ent.clave_e AS clave_e,boleto_ent.clave_c AS clave_cliente,cliente.nombre_c AS nombre_cliente,boleto_ent.clave_o AS clave_operador,operador.nombre_o AS nombre_operador," _
                & "boleto_ent.clave_p AS clave_producto,producto.nombre_p AS nombre_producto,boleto_ent.completo AS completo,boleto_ent.fecha_e AS fecha_entrada,boleto_ent.hora_e AS hora_entrada,boleto_ent.placas AS placas," _
                & "boleto_ent.conductor AS conductor,boleto_ent.peso_e AS peso_entrada,boleto_ent.observa_e AS observaciones_entrada,boleto_ent.unidad_e AS unidad_entrada," _
                & "boleto_ent.bascula_e AS bascula_entrada,boleto_ent.t_entrada AS t_entrada,boleto_ent.clave_u AS clave_u,boleto_ent.tipo_pesada AS tipo_pesada,boleto_ent.adicional1 AS adicional," _
                & "boleto_ent.precio_producto AS precio_producto,boleto_ent.tipo_precio AS tipo_precio,boleto_sal.fecha_s AS fecha_salida,boleto_sal.hora_s AS hora_salida," _
                & "boleto_sal.peso_s AS peso_salida,boleto_sal.peso_n AS peso_neto,boleto_sal.turno_s AS turno_salida,boleto_sal.bascula_s AS bascula_salida,boleto_sal.s_manual AS salida_manual," _
                & "boleto_sal.precio_total AS precio_total,boleto_sal.nombre_os AS nombre_os,boleto_sal.observa_s AS observa_salida " _
             & "From producto,operador ,cliente ,boleto_ent ,boleto_sal " _
             & "Where " _
                & "boleto_ent.clave_e = boleto_sal.clave_e " _
                & "AND boleto_ent.clave_c = cliente.clave_c " _
                & "AND boleto_ent.clave_o = operador.clave_o " _
                & "AND boleto_ent.clave_p = producto.clave_p " _
                & "AND cliente.nombre_c = 'AGROPECUARIA NUEVO SIGLO S.A.de C.V.' " _
                & "ORDER BY boleto_ent.clave_e DESC "
        
        ' Nuevo recordset ADO
        ' Abre el recordset
           
            If Rst.State = 1 Then
            
                Rst.Close
                
            End If
            
            Rst.Open C1, cnn, adOpenDynamic, adLockReadOnly
                   
            Rst.Requery
            
            If Rst2.State = 1 Then
            
                Rst2.Close
                
            End If
            
            Rst2.Open C2, cnn, adOpenDynamic, adLockReadOnly
                   
            Rst2.Requery
        
        Label1.Caption = Rst.RecordCount
        
        Label2.Caption = Label1.Caption
        
        Set ObjCarpeta = CreateObject("Scripting.FileSystemObject")
        Set Carpeta = ObjCarpeta.GetFolder("\\10.1.2.252\Embarques\pesadas\rev")
        
        Set ObjCarpeta2 = CreateObject("Scripting.FileSystemObject")
        Set Carpeta2 = ObjCarpeta.GetFolder("\\10.1.2.252\Embarques\pesadas\in")
        
        varcount = Carpeta.Files.Count 'Contador archivos in
        Label3.Caption = varcount
        
        varcount2 = Carpeta2.Files.Count 'Contador archivos out
        Label4.Caption = varcount2
            
        'varCuenta = (Label1.Caption + 1) - varcount
        Label5.Caption = varCuenta
    
End Sub

Private Sub Timer1_Timer()

    On Error Resume Next

    ' ----------------------------------------------------------------------------------------------
    ' -- Si el campo es nulo ( binario, o tipo desconocido etc..) devuelve False para no añadir el dato
    ' ----------------------------------------------------------------------------------------------

    With Rst
            
        If .State = 1 Then .Close
            
            .Open C1, cnn, adOpenStatic, adLockReadOnly
                   
            .Requery
                
    End With
    
    Label1.Caption = Rst.RecordCount

    If Label2.Caption <> Label1.Caption Then
    
        With Rst2
            
            If .State = 1 Then .Close
                
                .Open C2, cnn, adOpenStatic, adLockReadOnly
                       
                .Requery
                
        End With
        
        ' Llama a la función que genera el Csv con los datos del recordset
        R = Recordset_a_CSV(Rst2, "\\10.1.2.252\Embarques\pesadas\rev\Emb" + Label1.Caption + ".txt")
        
        Label2.Caption = Label1.Caption
        
    End If
    
End Sub
    
Private Sub Timer2_Timer()

    On Error Resume Next

    varcount = Carpeta.Files.Count 'Contador archivos in
    Label3.Caption = varcount
    
    If varcount > 0 Then
        
        varcount2 = Carpeta2.Files.Count 'Contador archivos out
        Label4.Caption = varcount2
        
        If varcount2 = 0 Then
        
            varCuenta = (Label1.Caption + 1) - varcount
            Label5.Caption = varCuenta
            FileCopy "\\10.1.2.252\Embarques\pesadas\rev\Emb" + Label5.Caption + ".txt", "\\10.1.2.252\Embarques\pesadas\in\Emb.txt" 'Copio archivo de rev a in
            FileCopy "\\10.1.2.252\Embarques\pesadas\rev\Emb" + Label5.Caption + ".txt", "\\10.1.2.252\Embarques\pesadas\hst\Emb" + Label5.Caption + ".txt" 'Copio archivo de rev a hst
            
            'Start append text to file
            Dim Canal%, i%
            Dim Dato(1) As String
            
            Dato(1) = "Emb" + Label5.Caption + ".txt"
            Canal = FreeFile
            Open "\\10.1.2.252\Embarques\pesadas\hst\History.txt" For Append As Canal
            Write #Canal, Dato(1)
            Close Canal
            
            Kill ("\\10.1.2.252\Embarques\pesadas\rev\Emb" + Label5.Caption + ".txt") 'Elimino archivo de rev
            
            varcount = Carpeta.Files.Count 'Contador archivos in
            Label3.Caption = varcount
            
            varcount2 = Carpeta2.Files.Count 'Contador archivos out
            Label4.Caption = varcount2
            
            varCuenta = (Label1.Caption + 1) - varcount
            Label5.Caption = varCuenta
        
        Else
        
            varcount = Carpeta.Files.Count 'Contador archivos in
            Label3.Caption = varcount
            
            varcount2 = Carpeta2.Files.Count 'Contador archivos out
            Label4.Caption = varcount2
            
            varCuenta = (Label1.Caption + 1) - varcount
            Label5.Caption = varCuenta
            
        End If
    
    Else
        
        varcount = Carpeta.Files.Count 'Contador archivos in
        Label3.Caption = varcount
            
        varcount2 = Carpeta2.Files.Count 'Contador archivos out
        Label4.Caption = varcount2
           
        varCuenta = (Label1.Caption + 1) - varcount
        Label5.Caption = varCuenta
        
    End If
    
End Sub
