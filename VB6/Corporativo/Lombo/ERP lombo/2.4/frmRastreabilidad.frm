VERSION 5.00
Begin VB.Form frmRastreabilidad 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecutar Reporte"
   ClientHeight    =   9075
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   17415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17175
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3135
         Index           =   1
         Left            =   4740
         TabIndex        =   1
         Top             =   2970
         Width           =   7935
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Index           =   2
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1440
            Width           =   6015
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "EXCEL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "PANTALLA"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Index           =   1
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   840
            Width           =   2415
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   465
            Index           =   0
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ARTICULO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "FOLIO"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   6
            Left            =   480
            TabIndex        =   3
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ORIGEN"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmRastreabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmRastreabilidad
'Proposito:     Rastreabilidad de Articulos mediante lote
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    '//RECORDSET
    Dim Rs          As New adodb.Recordset
    Dim RS1         As New adodb.Recordset
    Dim Rsreporte1  As New adodb.Recordset
    '//OTROS
    Dim i           As Long

    Private Sub Form_Load()
        On Error GoTo errHandler
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Combo1(0)
            .AddItem "Compra"
            .AddItem "Venta"
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_Click(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                With Combo1(1)
                    .Clear
                End With
                
                With Combo1(2)
                    .Clear
                End With
                
                With Combo1(0)
                    If .Text = "Compra" Then
                        With Rs
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select folio from PO_HEADERS_ALL_P order by CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT) desc;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            If .RecordCount > 0 Then
                                .MoveFirst
                                While Not .EOF
                                    Combo1(1).AddItem .Fields(0).Value
                                    .MoveNext
                                Wend
                            End If
                        End With
                    Else
                        With Rs
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select folio from PO_HEADERS_ALL_R order by CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT) desc;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            If .RecordCount > 0 Then
                                .MoveFirst
                                While Not .EOF
                                    Combo1(1).AddItem .Fields(0).Value
                                    .MoveNext
                                Wend
                            End If
                        End With
                    End If
                End With
            Case 1
                With Combo1(2)
                    .Clear
                End With
                
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select DISTINCT DescripcionArticulo from PO_LINES_ALL WHERE folio = '" & Combo1(1).Text & "' order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount > 0 Then
                        .MoveFirst
                        While Not .EOF
                            Combo1(2).AddItem .Fields(0).Value
                            .MoveNext
                        Wend
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click()
        On Error GoTo errHandler
        With Combo1(1)
            If .Text <> "" Then
                With Rsreporte1
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "SELECT CASE WHEN t2.folio LIKE 'C-%' THEN '1. Compras' WHEN t2.folio LIKE 'P-%' THEN '2. Produccion' WHEN t2.folio LIKE 'V-%' THEN '3. Ventas' END tipo, " & _
                          "       t2.folio, " & _
                          "       t2.lote, " & _
                          "       convert(varchar,t2.fecha,105)AS fecha, " & _
                          "       dbo.initcap((SELECT t1.nombre FROM po_headers_all_v t1 WHERE t2.folio = t1.folio))AS nombre, " & _
                          "       upper(t2.itemcodigo)AS itemcodigo, " & _
                          "       dbo.initcap(t2.itemdescricion)AS itemdescricion, " & _
                          "       t2.tipotransaccion, " & _
                          "       t2.cantidad, " & _
                          "       t2.udm " & _
                          "FROM mtl_material_transactions t2 " & _
                          "WHERE t2.lote IN( " & _
                          "        SELECT DISTINCT lote FROM mtl_material_transactions WHERE cancelado = 'No' " & _
                          "            AND folio IN(SELECT DISTINCT folio FROM mtl_material_transactions WHERE cancelado = 'No' " & _
                          "                    AND lote IN(SELECT DISTINCT lote FROM mtl_material_transactions WHERE folio = '" & Combo1(1).Text & "' and itemdescricion = '" & Combo1(2).Text & "'))) " & _
                          "  AND cancelado = 'No' " & _
                          "ORDER BY 1,5,CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT),3,7;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount <> 0 Then
                        Unload dsrRastreabilidad
                        With dsrRastreabilidad
                            Set .DataSource = Rsreporte1
                            With .Sections("Section4")
                                With .Controls("Label7")
                                    .Caption = Combo1(0).Text & " Folio: " & Combo1(1).Text
                                End With
                            End With
                            
                            With .Sections("Section1")
                                With .Controls("Text1")
                                    .DataField = "tipo"
                                End With
                                
                                With .Controls("Text2")
                                    .DataField = "nombre"
                                End With
                                
                                With .Controls("Text3")
                                    .DataField = "folio"
                                End With
                                
                                With .Controls("Text4")
                                    .DataField = "fecha"
                                End With
                                
                                With .Controls("Text5")
                                    .DataField = "itemdescricion"
                                End With
                                
                                With .Controls("Text6")
                                    .DataField = "tipotransaccion"
                                End With
                                
                                With .Controls("Text7")
                                    .DataField = "cantidad"
                                End With
                                
                                With .Controls("Text8")
                                    .DataField = "udm"
                                End With
                                
                                With .Controls("Text9")
                                    .DataField = "lote"
                                End With
                            End With
                            .Show 1
                        End With
                    End If
                End With
                Unload Me
            Else
                MsgBox "Seleccionar un folio", vbCritical, "Advertencia"
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        With Combo1(1)
            If .Text <> "" Then
                With Rsreporte1
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "SELECT CASE WHEN t2.folio LIKE 'C-%' THEN '1. Compras' WHEN t2.folio LIKE 'P-%' THEN '2. Produccion' WHEN t2.folio LIKE 'V-%' THEN '3. Ventas' END tipo, " & _
                          "       t2.folio, " & _
                          "       t2.lote, " & _
                          "       convert(varchar,t2.fecha,105)AS fecha, " & _
                          "       dbo.initcap((SELECT t1.nombre FROM po_headers_all_v t1 WHERE t2.folio = t1.folio))AS nombre, " & _
                          "       upper(t2.itemcodigo)AS itemcodigo, " & _
                          "       dbo.initcap(t2.itemdescricion)AS itemdescricion, " & _
                          "       t2.tipotransaccion, " & _
                          "       t2.cantidad, " & _
                          "       t2.udm " & _
                          "FROM mtl_material_transactions t2 " & _
                          "WHERE t2.lote IN( " & _
                          "        SELECT DISTINCT lote FROM mtl_material_transactions WHERE cancelado = 'No' " & _
                          "            AND folio IN(SELECT DISTINCT folio FROM mtl_material_transactions WHERE cancelado = 'No' " & _
                          "                    AND lote IN(SELECT DISTINCT lote FROM mtl_material_transactions WHERE folio = '" & Combo1(1).Text & "' and itemdescricion = '" & Combo1(2).Text & "'))) " & _
                          "  AND cancelado = 'No' " & _
                          "ORDER BY 1,5,CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT),3,7;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount <> 0 Then
                        'PARA EXPORTAR A EXCEL
                        Dim N As Long, sTemp As String
                        Dim FileName As String
                        FileName = App.Path & "\Temp\TEMP_RASTEABILIDAD_" & CStr(Format(Date, "YYYYMMDD")) & "_" & CStr(Format(Time, "HHMMSS")) & ".xls"
                        Open FileName For Output As #1
                        'ENCABEZADO
                        sTemp = "INFORME DE RASTREABILIDAD"
                        Print #1, sTemp
                        sTemp = vbNullString
                        With Combo1(0)
                            sTemp = "Origen: " & .Text
                        End With
                        
                        Print #1, sTemp
                        sTemp = vbNullString
                        With Combo1(1)
                            sTemp = "Folio: " & .Text
                        End With
                        
                        Print #1, sTemp
                        sTemp = vbNullString
                        sTemp = "Fecha de Ejecucion del informe: " & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
                        Print #1, sTemp
                        sTemp = vbNullString
                        Print #1, sTemp
                        sTemp = vbNullString
                        'CABECERA
                        For N = 0 To Rsreporte1.Fields.Count - 1
                            sTemp = sTemp & UCase(.Fields(N).Name) & IIf(N = .Fields.Count - 1, vbNullString, vbTab)
                        Next N
                        
                        Print #1, sTemp
                        sTemp = vbNullString
                        'DETALLE
                        .MoveFirst
                        Do Until .EOF
                            For N = 0 To .Fields.Count - 1
                                If N = 8 Then 'CONVERTIR A NUMERO
                                    sTemp = sTemp & Replace(CStr(.Fields(N).Value), ",", ".") & IIf(N = .Fields.Count - 1, vbNullString, vbTab)
                                Else
                                    sTemp = sTemp & .Fields(N).Value & IIf(N = .Fields.Count - 1, vbNullString, vbTab)
                                End If
                            Next N
                            
                            Print #1, sTemp
                            sTemp = vbNullString
                            .MoveNext
                        Loop
                        
                        Close #1
                        'PARA ABRIR EL ARCHIVO DE EXCEL AL TERMINAR DE EXPORTAR
                        Dim xltmp As Excel.Application
                        
                        Set xltmp = New Excel.Application
                        
                        With xltmp
                            With .Workbooks
                                .Open FileName
                            End With
                            
                            With .Range("A6", "J6")
                                With .Interior
                                    .Color = RGB(80, 80, 80)
                                End With
                                
                                With .Font
                                    .Color = RGB(255, 255, 255)
                                End With
                            End With
                            
                            With .ActiveWorkbook
                                .Save
                            End With
                            .Visible = True
                        End With
                    End If
                End With
                Unload Me
            Else
                MsgBox "Seleccionar un folio", vbCritical, "Advertencia"
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        With Rs
            If .State = 1 Then .Close
        End With
        
        With RS1
            If .State = 1 Then .Close
        End With
        
        With Rsreporte1
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set RS1 = Nothing
        Set Rsreporte1 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
