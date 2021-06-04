VERSION 5.00
Begin VB.Form frmBuscadorPedidos 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Pedidos"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   9015
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17220
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   8775
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   17055
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ACEPTAR"
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
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   8040
            Width           =   1455
         End
         Begin VB.ListBox List1 
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
            Height          =   6960
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   16695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Height          =   420
            Left            =   1320
            TabIndex        =   2
            Top             =   120
            Width           =   15495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "BUSCAR"
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
            Index           =   2
            Left            =   -960
            TabIndex        =   3
            Top             =   120
            Width           =   2055
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmBuscadorPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmBuscadorPedidos
'Proposito:     Buscar pedidos para importarlos en una venta
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        18/05/2021     Alfredo Hernandez    Se agrego validacion para inv.
'                                               negativos
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    Dim Rs As New adodb.Recordset
    Dim RS1 As New adodb.Recordset
    Dim Str As String
    Dim ArrStr() As String
    Dim viid                As String
    Dim videscripcion       As String
    Dim vicantidad          As String
    Dim viprecio            As String
    Dim i                   As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim c3                  As Long
    Dim c4                  As Long
    Dim viva                As String
    Dim listSubtotal        As String
    Dim listIva             As String
    Dim listTotal           As String
    Dim vLstCantidad        As String
    Dim vLstPrecio          As String
    
    Sub Form_Load()
        On Error GoTo errHandler
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "SELECT Distinct Folio, '|', Nombre From PO_LINES_ALL Where Tipo= 'Pedidos' AND cancelado= 'No' order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount <> 0 Then
                .MoveFirst
                With List1
                    .Clear
                End With
                
                While Not .EOF
                    List1.AddItem .Fields(0).Value & .Fields(1).Value & .Fields(2).Value
                    .MoveNext
                Wend
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorPedidos:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Text1_Change()
        On Error GoTo errHandler
        With List1
            .Clear
        End With
        With Rs
            If Text1 = "" Then
                .Filter = ""
                .Requery
                If .RecordCount <> 0 Then
                    .MoveFirst
                    While Not .EOF
                        List1.AddItem .Fields(1).Value
                        .MoveNext
                    Wend
                End If
            Else
                .Filter = "nombre like '*" & Text1 & "*' or folio = '" & Text1 & "'"
                .Requery
                If .RecordCount <> 0 Then
                    .MoveFirst
                    While Not .EOF
                        List1.AddItem .Fields(0).Value & .Fields(1).Value & .Fields(2).Value
                        .MoveNext
                    Wend
                End If
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorPedidos:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        With List1
            If .Text = "" Then
                MsgBox "Seleccione algún pedido", vbOKOnly, "Información"
            Else
                Str = .Text
                ArrStr() = Split(Str, "|")
                With RS1
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "SELECT * From PO_LINES_ALL Where folio = '" & ArrStr(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        With List1
                            .Clear
                        End With
                        With frmVentas
                            With .List1
                                .Clear
                            End With
                            
                            With .Text1(7)
                                .Text = RS1.Fields(3).Value
                            End With
                            
                            With .Combo1(0)
                                .Text = RS1.Fields(2).Value
                                .Enabled = False
                            End With
                            
                            With .Command1(2)
                                .Enabled = False
                            End With
                            
                            With .Text1(3)
                                .Text = RS1.Fields(18).Value
                            End With
                            
                            With .Combo1(2)
                                .Text = RS1.Fields(4).Value
                            End With
                        End With
                        
                        While Not .EOF
                            viid = .Fields(7).Value
                            videscripcion = Mid(.Fields(9).Value, 1, 47)
                            vicantidad = Replace(Format(Val(.Fields(10).Value), "0.00"), ",", ".")
                            viprecio = Replace(Format(Val(.Fields(12).Value), "0.00"), ",", ".")
                            If Get_ItemUDM(.Fields(7).Value) <> "Servicio" And Get_ItemCategoria(.Fields(7).Value) = "Inventario" Then
                                If PcInventarios = False Then
                                    If Val(Get_CantidadItem(.Fields(7).Value)) < Val(vicantidad) Then
                                        MsgBox "Existencia insuficiente, no se puede agregar " & videscripcion & " a la venta", vbCritical, "Advertencia"
                                        GoTo Siguiente
                                    End If
                                End If
                            End If
                            ' 1 - 10
                            c1 = 10 - Len(viid)
                            For i = 1 To c1
                                viid = " " & viid
                            Next i
                            ' 12 - 58
                            c2 = 47 - Len(videscripcion)
                            For i = 1 To c2
                                videscripcion = videscripcion & " "
                            Next i
                            ' 60 - 74
                            c3 = 15 - Len(vicantidad)
                            For i = 1 To c3
                                vicantidad = " " & vicantidad
                            Next i
                            ' 76 - 90
                            c4 = 15 - Len(viprecio)
                            For i = 1 To c4
                                viprecio = " " & viprecio
                            Next i
                            
                            With frmVentas
                                With .List1
                                    .AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                                End With
                            End With
Siguiente:
                            .MoveNext
                        Wend
                        listSubtotal = 0
                        listIva = 0
                        With frmVentas
                            With .List1
                                For i = 0 To .ListCount - 1
                                    .ListIndex = i
                                    vLstCantidad = Trim(Mid(.Text, 60, 15))
                                    vLstPrecio = Trim(Mid(.Text, 76, 15))
                                    viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                                Next i
                            End With
                            listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                            listIva = Replace(Format(listIva, "0.00"), ",", ".")
                            listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                            listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                            With .Text1(6)
                                .Text = listSubtotal
                            End With
                            
                            With .Text1(5)
                                .Text = listIva
                            End With
                            
                            With .Text1(4)
                                .Text = listTotal
                            End With
                        End With
                    Else
                        Unload Me
                    End If
                End With
                Unload Me
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorPedidos:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub List1_DblClick()
        On Error GoTo errHandler
        With List1
            If .Text = "" Then
                MsgBox "Seleccione algún pedido", vbOKOnly, "Información"
            Else
                Str = .Text
                ArrStr() = Split(Str, "|")
                With RS1
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "SELECT * From PO_LINES_ALL Where folio = '" & ArrStr(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        With List1
                            .Clear
                        End With
                        
                        With frmVentas
                            With .List1
                                .Clear
                            End With
                            
                            With .Text1(7)
                                .Text = RS1.Fields(3).Value
                            End With
                            
                            With .Combo1(0)
                                .Text = RS1.Fields(2).Value
                                .Enabled = False
                            End With
                            
                            With .Command1(2)
                                .Enabled = False
                            End With
                            
                            With .Text1(3)
                                .Text = RS1.Fields(18).Value
                            End With
                            
                            With .Combo1(2)
                                .Text = RS1.Fields(4).Value
                            End With
                        End With
                        
                        While Not .EOF
                            viid = .Fields(7).Value
                            videscripcion = Mid(.Fields(9).Value, 1, 47)
                            vicantidad = Replace(Format(Val(.Fields(10).Value), "0.00"), ",", ".")
                            viprecio = Replace(Format(Val(.Fields(12).Value), "0.00"), ",", ".")
                            If Get_ItemUDM(.Fields(7).Value) <> "Servicio" And Get_ItemCategoria(.Fields(7).Value) = "Inventario" Then
                                If PcInventarios = False Then
                                    If Val(Get_CantidadItem(.Fields(7).Value)) < Val(vicantidad) Then
                                        MsgBox "Existencia insuficiente, no se puede agregar " & videscripcion & " a la venta", vbCritical, "Advertencia"
                                        GoTo Siguiente
                                    End If
                                End If
                            End If
                            ' 1 - 10
                            c1 = 10 - Len(viid)
                            For i = 1 To c1
                                viid = " " & viid
                            Next i
                            ' 12 - 58
                            c2 = 47 - Len(videscripcion)
                            For i = 1 To c2
                                videscripcion = videscripcion & " "
                            Next i
                            ' 60 - 74
                            c3 = 15 - Len(vicantidad)
                            For i = 1 To c3
                                vicantidad = " " & vicantidad
                            Next i
                            ' 76 - 90
                            c4 = 15 - Len(viprecio)
                            For i = 1 To c4
                                viprecio = " " & viprecio
                            Next i
                            
                            With frmVentas
                                With .List1
                                    .AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                                End With
                            End With
Siguiente:
                            .MoveNext
                        Wend
                        listSubtotal = 0
                        listIva = 0
                        With frmVentas
                            With .List1
                                For i = 0 To .ListCount - 1
                                    .ListIndex = i
                                    vLstCantidad = Trim(Mid(.Text, 60, 15))
                                    vLstPrecio = Trim(Mid(.Text, 76, 15))
                                    viva = Get_ItemIva(Trim(Mid(.Text, 1, 10)))
                                    listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                                    listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                                Next i
                            End With
                            listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                            listIva = Replace(Format(listIva, "0.00"), ",", ".")
                            listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                            listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                            With .Text1(6)
                                .Text = listSubtotal
                            End With
                            
                            With .Text1(5)
                                .Text = listIva
                            End With
                            
                            With .Text1(4)
                                .Text = listTotal
                            End With
                        End With
                    Else
                        Unload Me
                    End If
                End With
                Unload Me
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorPedidos:List1_DblClick" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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
        
        Set Rs = Nothing
        Set RS1 = Nothing
        Set frmBuscadorPedidos = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorPedidos:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
