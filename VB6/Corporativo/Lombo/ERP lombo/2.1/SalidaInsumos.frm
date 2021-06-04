VERSION 5.00
Begin VB.Form frmSalidaInsumos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Insumos"
   ClientHeight    =   7185
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   10335
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
   ScaleHeight     =   7185
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   9855
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4020
            Left            =   240
            TabIndex        =   4
            Top             =   2520
            Width           =   9375
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "SalidaInsumos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   1800
            Picture         =   "SalidaInsumos.frx":0834
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1560
            TabIndex        =   1
            Top             =   720
            Width           =   4455
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1560
            TabIndex        =   0
            Text            =   "Combo1"
            Top             =   240
            Width           =   8055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Id                Descripción                                                                  Cantidad"
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
            Index           =   1
            Left            =   480
            TabIndex        =   9
            Top             =   2160
            Width           =   9255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Producto"
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
            Index           =   0
            Left            =   -600
            TabIndex        =   8
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
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
            Index           =   5
            Left            =   -360
            TabIndex        =   7
            Top             =   840
            Width           =   1815
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmSalidaInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset
    Dim Rs1                 As New adodb.Recordset
    
    '//ARTICULO
    Dim vItemId             As Long
    Dim viid                As String
    Dim videscripcion       As String
    Dim vicantidad          As String
    
    '//OTROS
    Dim St                  As String
    Dim i                   As Long
    Dim X                   As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim c3                  As Long
    
    '//VALORES PARA INSERTAR
    Dim v1                  As String
    Dim v4                  As Long
    Dim v5                  As String
    Dim v6                  As String
    Dim v7                  As String
    
    '//LOTE
    Dim ControlLote         As Long
    Dim InLoteExiste        As Long
    Dim CantidadRestante    As String
    Dim vLote               As String
    Dim vCantidadLote       As String
    Dim vCurrentLote        As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
    
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_SYSTEM_ITEMS t1 where UDM <> 'Servicio' and Categoria = 'Inventario' order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            
            If .RecordCount > 0 Then
                Combo1(0).Clear
                
                While Not .EOF
                    Combo1(0).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                    
                    .MoveNext
                Wend
            End If
        End With
        
        If Rs1.RecordCount > 0 Then
            
            With Text1(0)
                .BackColor = COLOR_NO_ENCONTRADO
                .Text = ""
            End With
            
            Combo1(0).BackColor = COLOR_NO_ENCONTRADO
         
            List1.Clear
            Combo1(0).Text = ""
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"

            Unload Me
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                With Combo1(0)
                    If .Text <> "" Then
                        .BackColor = COLOR_NORMAL
                        vItemId = Get_ItemId(Combo1(0).Text)
                        
                        With Rs1
                            .Filter = ""
                            .Filter = "Id = " & vItemId
                            .Requery
                        End With
                    Else
                        .BackColor = COLOR_NO_ENCONTRADO
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        
        On Error GoTo errHandler
        
        Static cadena As String
        
        Select Case Index
            Case 0
                With Combo1(0)
                    ' si pesionamos las teclas de las flechas sale de la rutina
                    If KeyCode = vbKeyUp Then Exit Sub
                    If KeyCode = vbKeyDown Then Exit Sub
                    If KeyCode = vbKeyLeft Then Exit Sub
                    If KeyCode = vbKeyRight Then Exit Sub
                    
                    ' verifica qu no se presionó la tecla backspace
                    If KeyCode <> vbKeyBack Then
                        cadena = Mid(.Text, 1, Len(.Text) - .SelLength)
                    Else
                        '...tecla backspace
                        If cadena <> "" Then
                            cadena = Mid(cadena, 1, Len(cadena) - 1)
                        End If
                    End If
                    
                    For i = 0 To .ListCount - 1
                        If UCase(cadena) = UCase(Mid(.List(i), 1, Len(cadena))) Then
                            .ListIndex = i
                            
                            Exit For
                        End If
                    Next
                    
                    ' Seelecciona
                    .SelStart = Len(cadena)
                    .SelLength = Len(.Text)
                    
                    If .ListIndex = -1 Then
                        ' color de fondo del combo en caso de que no hay coincidencias
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        ' Backcolor normal cuando hay coincidencia
                        .BackColor = COLOR_NORMAL
                        
                        vItemId = Get_ItemId(Combo1(0).Text)
                        
                        With Rs1
                            .Filter = ""
                            .Filter = "Id = " & vItemId
                            .Requery
                        End With
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Text1_Change(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                With Text1(0)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                If Val(Text1(0)) < 0 Then
                    MsgBox "la cantidad es inválida", vbCritical, "Error"
                    
                    Exit Sub
                End If
                
                If Combo1(0) <> "" And Text1(0) <> "" Then
                    viid = Rs1.Fields(0).Value
                    videscripcion = Mid(Rs1.Fields(2).Value, 1, 43)
                    vicantidad = Replace(Format(Val(Text1(0)) * -1, "0.00"), ",", ".")
                    
                    ' 1 - 11
                    c1 = 10 - Len(viid)
                    
                    For i = 1 To c1
                        viid = " " & viid
                    Next i
                    
                    ' 12 - 55
                    c2 = 43 - Len(videscripcion)
                    
                    For i = 1 To c2
                        videscripcion = videscripcion & " "
                    Next i
                    
                    ' 56 - 64
                    c3 = 11 - Len(vicantidad)
                    
                    For i = 1 To c3
                        vicantidad = " " & vicantidad
                    Next i
                    
                    For X = 0 To List1.ListCount - 1
                        If UCase(Trim(Mid(List1.List(X), 1, 54))) = UCase(Trim(viid & " " & videscripcion)) Then
                            MsgBox "El articulo ya esta en la lista", vbOKOnly, "Atención"
                            
                            Exit Sub
                        End If
                    Next
                    
                    List1.AddItem viid & " " & videscripcion & " " & vicantidad
                    
                    Text1(0) = ""
                    Text1(0).BackColor = COLOR_NO_ENCONTRADO
                    
                    With Combo1(0)
                        .Text = ""
                        .BackColor = COLOR_NO_ENCONTRADO
                        .SetFocus
                    End With
                Else
                    MsgBox "Llenar todos los campos", vbCritical, "Error"
                    
                    Text1(0).SetFocus
                End If
            
            Case 1
                
                'Variables
                Dim intX As Long
                
                With List1
                    intX = .ListIndex
                    
                    .RemoveItem intX
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            
            For i = 0 To List1.ListCount - 1
                List1.ListIndex = i
                List1.SetFocus
                
                v4 = Trim(Mid(List1.Text, 1, 10))
                v1 = Get_ItemUDM(v4)
                v5 = Get_ItemCod(v4)
                v6 = Get_ItemDesc(v4)
                v7 = Replace(Trim(Mid(List1.Text, 56, 11)), ",", ".")
                
                'lote
                ControlLote = Get_ItemLote(v4)
                            
                'si tiene control de lote
                If ControlLote = 1 Then
                    CantidadRestante = v7
                
                    'mientras no se complete la cantidad necesaria
                    While Val(CantidadRestante) < 0
                        'obtenemos lote mas antiguo y existencia de ese lote
                        vLote = ""
                        vLote = Get_LoteConsumo(v4)
                        vCantidadLote = Get_LoteConsumoCantidad(v4)
                                    
                        'si existe algun lote
                        If vLote <> "" Then
                        
                            .AddNew
                                .Fields(1) = v4
                                .Fields(2) = v5
                                .Fields(3) = v6
                                .Fields(4) = Date
                                .Fields(5) = "Salida de Insumos"
                                .Fields(7) = v1
                                .Fields(9) = "No"
                                
                                If Val(vCantidadLote) > Val(CantidadRestante) Then
                                    .Fields(10) = vLote 'lote
                                    .Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".") 'cantidad
                                            
                                    CantidadRestante = "0"
                                Else
                                    .Fields(10) = vLote 'lote
                                    .Fields(6) = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".") 'cantidad
                                            
                                    CantidadRestante = Val(CantidadRestante) - Val(vCantidadLote)
                                End If
                            .Update
                            .Requery
                        Else
                            MsgBox "No se puede hacer la salida  por " & CantidadRestante & " de " & v6 & ", no hay lotes existentes", vbCritical, "Error"
                                
                            CantidadRestante = "0"
                        End If
                    Wend
                Else
                    CantidadRestante = Get_CantidadItem(v4)
                    
                    If Val(CantidadRestante) >= Val(v7) * -1 Then
                        .AddNew
                            .Fields(1) = v4
                            .Fields(2) = v5
                            .Fields(3) = v6
                            .Fields(4) = Date
                            .Fields(5) = "Salida de Insumos"
                            .Fields(6) = v7
                            .Fields(7) = v1
                            .Fields(9) = "No"
                        .Update
                        .Requery
                    Else
                        MsgBox "No se puede hacer la salida  por " & v7 & " de " & v6 & ", no hay suficientes existencias", vbCritical, "Error"
                    End If
                    
                    CantidadRestante = "0"
                End If
            Next i
        End With
        
        Unload frmSalidaInsumos
        
        Set frmSalidaInsumos = Nothing
        
        frmSalidaInsumos.Show
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmSalidaInsumos:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
