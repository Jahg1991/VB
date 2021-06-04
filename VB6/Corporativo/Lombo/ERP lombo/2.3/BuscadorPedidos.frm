VERSION 5.00
Begin VB.Form frmBuscadorPedidos 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6660
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6375
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   2520
            Picture         =   "BuscadorPedidos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2400
            Width           =   1455
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   6135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1080
            TabIndex        =   2
            Top             =   120
            Width           =   5175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   -1080
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
    Option Explicit
    
    Dim Rs As New adodb.Recordset
    Dim Rs1 As New adodb.Recordset
    
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
                
                List1.Clear
                
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
        
        List1.Clear
        
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
        
        If List1.Text = "" Then
            MsgBox "Seleccione algún pedido", vbOKOnly, "Información"
        Else
            Str = List1.Text
            ArrStr() = Split(Str, "|")
            
            With Rs1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "SELECT * From PO_LINES_ALL Where folio = '" & ArrStr(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    List1.Clear
                    frmVentas.List1.Clear
                    
                    frmVentas.Text1(7).Text = Rs1.Fields(3).Value
                    frmVentas.Combo1(0).Text = Rs1.Fields(2).Value
                    
                    frmVentas.Combo1(0).Enabled = False
                    frmVentas.Command1(2).Enabled = False
                    
                    frmVentas.Text1(3).Text = Rs1.Fields(18).Value
                    frmVentas.Combo1(2).Text = Rs1.Fields(4).Value
                    
                    While Not .EOF
                        viid = .Fields(7).Value
                        videscripcion = Mid(.Fields(9).Value, 1, 47)
                        vicantidad = Replace(Format(Val(.Fields(10).Value), "0.00"), ",", ".")
                        viprecio = Replace(Format(Val(.Fields(12).Value), "0.00"), ",", ".")
                        
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
                        
                        frmVentas.List1.AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                        
                        .MoveNext
                    Wend
                    
                    listSubtotal = 0
                    listIva = 0
                    
                    For i = 0 To frmVentas.List1.ListCount - 1
                        frmVentas.List1.ListIndex = i
                        
                        vLstCantidad = Trim(Mid(frmVentas.List1.Text, 60, 15))
                        vLstPrecio = Trim(Mid(frmVentas.List1.Text, 76, 15))
                        viva = Get_ItemIva(Trim(Mid(frmVentas.List1.Text, 1, 10)))
                        listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                        listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                    Next i
                    
                    listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                    listIva = Replace(Format(listIva, "0.00"), ",", ".")
                    listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                    listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                    
                    frmVentas.Text1(6) = listSubtotal
                    frmVentas.Text1(5) = listIva
                    frmVentas.Text1(4) = listTotal
                Else
                    Unload Me
                End If
            End With
            
            Unload Me
        End If
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
        
        If List1.Text = "" Then
            MsgBox "Seleccione algún pedido", vbOKOnly, "Información"
        Else
            Str = List1.Text
            ArrStr() = Split(Str, "|")
            
            With Rs1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "SELECT * From PO_LINES_ALL Where folio = '" & ArrStr(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    List1.Clear
                    frmVentas.List1.Clear
                    
                    frmVentas.Text1(7).Text = Rs1.Fields(3).Value
                    frmVentas.Combo1(0).Text = Rs1.Fields(2).Value
                    
                    frmVentas.Combo1(0).Enabled = False
                    frmVentas.Command1(2).Enabled = False
                    
                    frmVentas.Text1(3).Text = Rs1.Fields(18).Value
                    frmVentas.Combo1(2).Text = Rs1.Fields(4).Value
                    
                    While Not .EOF
                        viid = .Fields(7).Value
                        videscripcion = Mid(.Fields(9).Value, 1, 47)
                        vicantidad = Replace(Format(Val(.Fields(10).Value), "0.00"), ",", ".")
                        viprecio = Replace(Format(Val(.Fields(12).Value), "0.00"), ",", ".")
                        
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
                        
                        frmVentas.List1.AddItem viid & " " & videscripcion & " " & vicantidad & " " & viprecio
                        
                        .MoveNext
                    Wend
                    
                    listSubtotal = 0
                    listIva = 0
                    
                    For i = 0 To frmVentas.List1.ListCount - 1
                        frmVentas.List1.ListIndex = i
                        
                        vLstCantidad = Trim(Mid(frmVentas.List1.Text, 60, 15))
                        vLstPrecio = Trim(Mid(frmVentas.List1.Text, 76, 15))
                        viva = Get_ItemIva(Trim(Mid(frmVentas.List1.Text, 1, 10)))
                        listSubtotal = (Val(vLstCantidad) * Val(vLstPrecio)) + listSubtotal
                        listIva = ((Val(vLstCantidad) * Val(vLstPrecio)) * Val(viva)) + listIva
                    Next i
                    
                    listTotal = Val(Replace(Format(listSubtotal, "0.00"), ",", ".")) + Val(Replace(Format(listIva, "0.00"), ",", "."))
                    listIva = Replace(Format(listIva, "0.00"), ",", ".")
                    listSubtotal = Replace(Format(listSubtotal, "0.00"), ",", ".")
                    listTotal = Replace(Format(listTotal, "0.00"), ",", ".")
                    
                    frmVentas.Text1(6) = listSubtotal
                    frmVentas.Text1(5) = listIva
                    frmVentas.Text1(4) = listTotal
                Else
                    Unload Me
                End If
            End With
            
            Unload Me
        End If
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
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        
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
