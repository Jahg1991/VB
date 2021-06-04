VERSION 5.00
Begin VB.Form frmAjusteInventario 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Ajuste de Inventario"
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
   Moveable        =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      ForeColor       =   &H00404040&
      Height          =   2295
      Index           =   0
      Left            =   5400
      TabIndex        =   16
      Top             =   3480
      Width           =   6780
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   6255
         Begin VB.CommandButton Command1 
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
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   960
            Width           =   1455
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CATEGORIA"
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
            Index           =   4
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   9015
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17295
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   8775
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   17175
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
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   5280
            Width           =   1455
         End
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
            Index           =   2
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   5280
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "AJUSTAR"
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
            TabIndex        =   20
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   13320
            TabIndex        =   9
            Top             =   3720
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   13320
            TabIndex        =   13
            Top             =   4680
            Width           =   3500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   5
            Left            =   8760
            TabIndex        =   12
            Top             =   4680
            Width           =   4455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   4320
            TabIndex        =   11
            Top             =   4680
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   4680
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   4320
            TabIndex        =   8
            Top             =   3720
            Width           =   8895
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   1920
            TabIndex        =   7
            Top             =   3720
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   6
            Top             =   3720
            Width           =   1695
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
            Height          =   2055
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   16695
         End
         Begin VB.ListBox List2 
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
            Height          =   2055
            Left            =   120
            TabIndex        =   2
            Top             =   6360
            Width           =   16695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAjusteInventario.frx":0000
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
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   6000
            Width           =   15735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAjusteInventario.frx":008D
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
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   15855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAjusteInventario.frx":0117
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
            TabIndex        =   5
            Top             =   4320
            Width           =   16695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAjusteInventario.frx":01C1
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
            Left            =   120
            TabIndex        =   4
            Top             =   3360
            Width           =   15975
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
Attribute VB_Name = "frmAjusteInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmAjusteInventario
'Proposito:     Ajustar existencias de articulos
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
    Dim Rs      As New adodb.Recordset
    Dim RS1     As New adodb.Recordset
    Dim Rs2     As New adodb.Recordset
    '//OTROS
    Dim i       As Long
    Dim c1      As String
    Dim c2      As String
    Dim c3      As String
    Dim c4      As String
    Dim c5      As String
    Dim c6      As String
    Dim nc      As Long
    Dim intX    As Long
    Dim X       As Long
    '//VALORES PARA INSERTAR
    Dim v1      As Long
    Dim v2      As String
    Dim v3      As String
    Dim v4      As String
    Dim v5      As String
    Dim v6      As String
    Dim v7      As String
    Dim v8      As String
    Dim v9      As String
    Dim v10     As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        With Frame1(1)
            .Enabled = False
        End With
        
        With Frame3(0)
            .Enabled = True
        End With
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With List1
            .Clear
        End With
        
        With List2
            .Clear
        End With
        
        With Rs2
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select distinct tipo from MTL_ON_HAND_QUANTITIES order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    Combo1.AddItem .Fields(0)
                    .MoveNext
                Wend
                .MoveFirst
                Combo1.Text = .Fields(0)
                .Close
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                With List1
                    If Mid(.Text, 1, 10) <> "" Then
                        For i = 0 To (7)
                            With Text1(i)
                                .Text = ""
                            End With
                        Next i
                        Text1(0) = Trim(Mid(.Text, 1, 10))
                        Text1(4) = Trim(Mid(.Text, 68, 14))
                        Text1(7) = Trim(Mid(.Text, 94, 19))
                    Else
                        MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
                    End If
                End With
            Case 1
                With Text1(0)
                    If .Text <> "" Then
                        With Text1(6)
                            If Val(.Text) <> 0 Then
                                With Text1(0)
                                    c1 = Mid(.Text, 1, 10)
                                End With
                                
                                With Text1(1)
                                    c2 = Mid(.Text, 1, 10)
                                End With
                                
                                With Text1(2)
                                    c3 = Mid(.Text, 1, 44)
                                End With
                                c4 = Mid(.Text, 1, 14)
                                With Text1(3)
                                    c5 = Mid(.Text, 1, 10)
                                End With
                                
                                With Text1(7)
                                    c6 = Mid(.Text, 1, 19)
                                End With
                                nc = 10 - Len(c1)
                                For i = 1 To nc
                                    c1 = " " & c1
                                Next i
                                nc = 10 - Len(c2)
                                For i = 1 To nc
                                    c2 = c2 & " "
                                Next i
                                nc = 44 - Len(c3)
                                For i = 1 To nc
                                    c3 = c3 & " "
                                Next i
                                nc = 14 - Len(c4)
                                For i = 1 To nc
                                    c4 = " " & c4
                                Next i
                                nc = 10 - Len(c5)
                                For i = 1 To nc
                                    c5 = c5 & " "
                                Next i
                                nc = 19 - Len(c6)
                                For i = 1 To nc
                                    c6 = c6 & " "
                                Next i
                                
                                With List1
                                    For X = 0 To .ListCount - 1
                                        If UCase(Trim(Mid(List2.List(X), 1, 10))) = UCase(Trim(Mid(Text1(0).Text, 1, 10))) Then
                                            If UCase(Trim(Mid(List2.List(X), 94, 19))) = UCase(Trim(Mid(Text1(7), 1, 19))) Then
                                                With List2
                                                    MsgBox UCase(Trim(Mid(.List(X), 94, 19)))
                                                End With
                                                
                                                With Text1(7)
                                                    MsgBox "El articulo " & UCase(Trim(Mid(.Text, 1, 19))) & " ya esta en la lista", vbOKOnly, "Atención"
                                                End With
                                                
                                                With Text1(i)
                                                    For i = 0 To 6
                                                        .Text = ""
                                                    Next i
                                                End With
                                                
                                                Exit Sub
                                            End If
                                        End If
                                    Next
                                End With
                                
                                With List2
                                    .AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5 & " " & c6
                                End With
                            End If
                        End With
                        
                        For i = 0 To 7
                            With Text1(i)
                                .Text = ""
                            End With
                        Next i
                    Else
                        MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
                    End If
                End With
            Case 2
                With List2
                    intX = .ListIndex
                    .RemoveItem intX
                End With
            Case 3
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from MTL_ON_HAND_QUANTITIES where tipo = '" & Combo1.Text & "' order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    If .RecordCount > 0 Then
                        Do Until .EOF
                            c1 = Mid(.Fields(1).Value, 1, 10)
                            c2 = Mid(.Fields(2).Value, 1, 10)
                            c3 = Mid(.Fields(3).Value, 1, 44)
                            c4 = Replace(Format(Mid(.Fields(4).Value, 1, 14), "0.00"), ",", ".")
                            c5 = Mid(.Fields(5).Value, 1, 10)
                            If IsNull(.Fields(6).Value) = False Then
                                c6 = Mid(.Fields(6).Value, 1, 19)
                            Else
                                c6 = ""
                            End If
                            nc = 10 - Len(c1)
                            For i = 1 To nc
                                c1 = " " & c1
                            Next i
                            nc = 10 - Len(c2)
                            For i = 1 To nc
                                c2 = c2 & " "
                            Next i
                            nc = 44 - Len(c3)
                            For i = 1 To nc
                                c3 = c3 & " "
                            Next i
                            nc = 14 - Len(c4)
                            For i = 1 To nc
                                c4 = " " & c4
                            Next i
                            nc = 10 - Len(c5)
                            For i = 1 To nc
                                c5 = c5 & " "
                            Next i
                            nc = 19 - Len(c6)
                            For i = 1 To nc
                                c6 = c6 & " "
                            Next i
                            
                            With List1
                                .AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5 & " " & c6
                            End With
                            .MoveNext
                        Loop
                    End If
                End With
            With Frame1(1)
                .Enabled = True
            End With
            
            With Frame3(0)
                .Visible = False
            End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Command1_Click" & vbTab & err.Number & vbTab & err.Description
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
                        For i = 1 To 6
                            With Text1(i)
                                .Text = ""
                            End With
                        Next i
                    Else
                        With Text1(0)
                            Text1(1) = Get_ItemCod(.Text)
                            Text1(2) = Get_ItemDesc(.Text)
                            Text1(3) = Get_ItemUDM(.Text)
                        End With
                    End If
                End With
            Case 5
                With Text1(5)
                    If .Text = "" Then
                        With Text1(6)
                            .Text = 0
                        End With
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        With Text1(6)
                            .Text = Replace(Format(Val(Text1(5)) - Val(Text1(4)), "0.00"), ",", ".")
                        End With
                        .BackColor = COLOR_NORMAL
                    End If
                End With
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
        If vbq = vbYes Then
            With RS1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Filter = ""
                .Requery
                With List2
                    If .ListCount > 0 Then
                        For i = 0 To .ListCount - 1
                            .ListIndex = i
                            .SetFocus
                            v1 = Trim(Mid(.Text, 1, 10))
                            v2 = Get_ItemCod(v1)
                            v3 = Get_ItemDesc(v1)
                            v4 = Date
                            v6 = Trim(Mid(List2.Text, 68, 14))
                            If Val(v6) < 0 Then
                                v5 = "Salida Ajuste"
                            Else
                                v5 = "Entrada Ajuste"
                            End If
                            v7 = Get_ItemUDM(v1)
                            v8 = "Ajuste " & Date
                            v9 = "No"
                            v10 = Trim(Mid(.Text, 94, 19))
                            With RS1
                                .AddNew
                                    With .Fields(1)
                                        .Value = v1         'item_id
                                    End With
                                    
                                    With .Fields(2)
                                        .Value = v2         'codigo
                                    End With
                                    
                                    With .Fields(3)
                                        .Value = v3        'descripcion
                                    End With
                                    
                                    With .Fields(4)
                                        .Value = v4         'fecha
                                    End With
                                    
                                    With .Fields(5)
                                        .Value = v5         'transaccion
                                    End With
                                    
                                    With .Fields(6)
                                        .Value = v6         'cantidad
                                    End With
                                    
                                    With .Fields(7)
                                        .Value = v7         'udm
                                    End With
                                    
                                    With .Fields(8)
                                        .Value = v8         'folio
                                    End With
                                    
                                    With .Fields(9)
                                        .Value = v9         'cancelado
                                    End With
                                    
                                    If v10 <> "" Then
                                        With .Fields(10)
                                            .Value = v10    'lote
                                        End With
                                    End If
                                .Update
                                .Requery
                            End With
                        Next i
                        MsgBox "Ajuste Finalizado", vbOKOnly, "Información"
                    Else
                        MsgBox "No se agregaron artículos para ajustar", vbCritical, "Advertencia"
                        Exit Sub
                    End If
                End With
            End With
            Unload Me
        Else
            Exit Sub
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        
        With Rs2
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set RS1 = Nothing
        Set Rs2 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
