VERSION 5.00
Begin VB.Form frmSalidaInsumos 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Insumos"
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   17175
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8655
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   16935
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ELIMINAR"
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
            Index           =   1
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "AÑADIR"
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
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ListBox List1 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   6045
            Left            =   120
            TabIndex        =   2
            Top             =   2520
            Width           =   16695
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Top             =   660
            Width           =   4455
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00808080&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
            Height          =   465
            Index           =   0
            Left            =   1680
            TabIndex        =   0
            Text            =   "Combo1"
            Top             =   120
            Width           =   15135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "ID               DESCRIPCION                                                              CANTIDAD"
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
            Left            =   240
            TabIndex        =   9
            Top             =   2160
            Width           =   14295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTO"
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
            Left            =   -600
            TabIndex        =   6
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD"
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
            Index           =   5
            Left            =   -360
            TabIndex        =   5
            Top             =   660
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
'***********************************************************************************
'Nombre:        frmSalida Insumos
'Proposito:     Registro de salida de insumos
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
Dim Rs As New adodb.Recordset
Dim RS1 As New adodb.Recordset
'//ARTICULO
Dim vItemId As Long
Dim viid As String
Dim videscripcion As String
Dim vicantidad As String
'//OTROS
Dim St As String
Dim i As Long
Dim X As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
'//VALORES PARA INSERTAR
Dim v1 As String
Dim v4 As Long
Dim v5 As String
Dim v6 As String
Dim v7 As String
'//LOTE
Dim ControlLote As Boolean
Dim InLoteExiste As Long
Dim CantidadRestante As String
Dim vLote As String
Dim vCantidadLote As String
Dim vCurrentLote As String

Private Sub Form_Load()
    On Error GoTo errHandler
    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With RS1
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from MTL_SYSTEM_ITEMS t1 where UDM <> 'Servicio' and Categoria = 'Inventario' order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Filter = ""
        .Requery
        If .RecordCount > 0 Then
            With Combo1(0)
                .Clear
            End With

            While Not .EOF
                Combo1(0).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                .MoveNext
            Wend
        End If

        If .RecordCount > 0 Then
            With Text1(0)
                .BackColor = COLOR_NO_ENCONTRADO
                .Text = ""
            End With

            With Combo1(0)
                .BackColor = COLOR_NO_ENCONTRADO
                .Text = ""
            End With

            With List1
                .Clear
            End With
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
        End If
    End With
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
                vItemId = Get_ItemId(.Text)
                With RS1
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
                'color de fondo del combo en caso de que no hay coincidencias
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                ' Backcolor normal cuando hay coincidencia
                .BackColor = COLOR_NORMAL
                vItemId = Get_ItemId(.Text)
                With RS1
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
        With Text1(0)
            If Val(.Text) < 0 Then
                MsgBox "la cantidad es inválida", vbCritical, "Error"
                Exit Sub
            End If

            If Combo1(0) <> "" And .Text <> "" Then
                With RS1
                    viid = .Fields(0).Value
                    videscripcion = Mid(.Fields(2).Value, 1, 43)
                End With
                vicantidad = Replace(Format(Val(.Text) * -1, "0.00"), ",", ".")
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

                With List1
                    For X = 0 To .ListCount - 1
                        If UCase(Trim(Mid(.List(X), 1, 54))) = UCase(Trim(viid & " " & videscripcion)) Then
                            MsgBox "El articulo ya esta en la lista", vbOKOnly, "Atención"
                            Exit Sub
                        End If
                    Next
                    .AddItem viid & " " & videscripcion & " " & vicantidad
                End With
                .Text = ""
                .BackColor = COLOR_NO_ENCONTRADO
                With Combo1(0)
                    .Text = ""
                    .BackColor = COLOR_NO_ENCONTRADO
                    .SetFocus
                End With
            Else
                MsgBox "Llenar todos los campos", vbCritical, "Error"
                .SetFocus
            End If
        End With
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
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            With List1
                For i = 0 To .ListCount - 1
                    .ListIndex = i
                    .SetFocus
                    v4 = Trim(Mid(.Text, 1, 10))
                    v1 = Get_ItemUDM(v4)
                    v5 = Get_ItemCod(v4)
                    v6 = Get_ItemDesc(v4)
                    v7 = Replace(Trim(Mid(.Text, 56, 11)), ",", ".")
                    'lote
                    ControlLote = Get_ItemLote(v4)
                    'si tiene control de lote
                    If ControlLote = True Then
                        CantidadRestante = v7
                        'mientras no se complete la cantidad necesaria
                        While Val(CantidadRestante) < 0
                            'obtenemos lote mas antiguo y existencia de ese lote
                            vLote = ""
                            vLote = Get_LoteConsumo(v4)
                            vCantidadLote = Get_LoteConsumoCantidad(v4)
                            'si existe algun lote
                            If vLote <> "" Then
                                With Rs
                                    .AddNew
                                    With .Fields(1)
                                        .Value = v4                                                                 'id
                                    End With

                                    With .Fields(2)
                                        .Value = v5                                                                 'codigo
                                    End With

                                    With .Fields(3)
                                        .Value = v6                                                                 'descripcion
                                    End With

                                    With .Fields(4)
                                        .Value = Date                                                               'fecha
                                    End With

                                    With .Fields(5)
                                        .Value = "Salida de Insumos"                                                'transaccion
                                    End With

                                    With .Fields(7)
                                        .Value = v1                                                                 'udm
                                    End With

                                    With .Fields(9)
                                        .Value = "No"                                                               'cancelado
                                    End With

                                    If Val(vCantidadLote) > Val(CantidadRestante) Then
                                        With .Fields(10)
                                            .Value = vLote                                                          'lote
                                        End With

                                        With .Fields(6)
                                            .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")  'cantidad
                                        End With
                                        CantidadRestante = "0"
                                    Else
                                        With .Fields(10)
                                            .Value = vLote                                                          'lote
                                        End With

                                        With .Fields(6)
                                            .Value = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".")     'cantidad
                                        End With
                                        CantidadRestante = Val(CantidadRestante) - Val(vCantidadLote)
                                    End If
                                    .Update
                                    .Requery
                                End With
                            Else
                                MsgBox "No se puede hacer la salida  por " & CantidadRestante & " de " & v6 & ", no hay lotes existentes", vbCritical, "Error"
                                CantidadRestante = "0"
                            End If
                        Wend
                    Else
                        CantidadRestante = Get_CantidadItem(v4)
                        If Val(CantidadRestante) >= Val(v7) * -1 Then
                            With Rs
                                .AddNew
                                With .Fields(1)
                                    .Value = v4                     'id
                                End With

                                With .Fields(2)
                                    .Value = v5                     'codigo
                                End With

                                With .Fields(3)
                                    .Value = v6                     'descripcion
                                End With

                                With .Fields(4)
                                    .Value = Date                   'fecha
                                End With

                                With .Fields(5)
                                    .Value = "Salida de Insumos"    'transaccion
                                End With

                                With .Fields(6)
                                    .Value = v7                     'cantidad
                                End With

                                With .Fields(7)
                                    .Value = v1                     'udm
                                End With

                                With .Fields(9)
                                    .Value = "No"                   'cancelado
                                End With
                                .Update
                                .Requery
                            End With
                        Else
                            MsgBox "No se puede hacer la salida  por " & v7 & " de " & v6 & ", no hay suficientes existencias", vbCritical, "Error"
                        End If
                        CantidadRestante = "0"
                    End If
                Next i
            End With
        End With
        Unload frmSalidaInsumos
        Set frmSalidaInsumos = Nothing

        With frmSalidaInsumos
            .Show
        End With
    End If
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
    With Rs
        If .State = 1 Then .Close
    End With

    With RS1
        If .State = 1 Then .Close
    End With

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set RS1 = Nothing
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
