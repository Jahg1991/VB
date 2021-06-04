VERSION 5.00
Begin VB.Form frmAjusteInventario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ajuste de Inventario"
   ClientHeight    =   7665
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   16725
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
   ScaleHeight     =   7665
   ScaleWidth      =   16725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   2055
      Index           =   0
      Left            =   5040
      TabIndex        =   19
      Top             =   2160
      Width           =   6540
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   6255
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   4575
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   3
            Left            =   2400
            Picture         =   "AjusteInventario.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16455
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   7215
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16215
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   7
            Left            =   13320
            TabIndex        =   12
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   6
            Left            =   9720
            TabIndex        =   16
            Top             =   3960
            Width           =   3500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   5
            Left            =   6120
            TabIndex        =   15
            Top             =   3960
            Width           =   3500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   4
            Left            =   2520
            TabIndex        =   14
            Top             =   3960
            Width           =   3500
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   3960
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   2
            Left            =   4320
            TabIndex        =   11
            Top             =   3120
            Width           =   8895
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   1
            Left            =   1920
            TabIndex        =   10
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   3120
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   2
            Left            =   1680
            Picture         =   "AjusteInventario.frx":080F
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   120
            Picture         =   "AjusteInventario.frx":10E0
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4440
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   120
            Picture         =   "AjusteInventario.frx":1914
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2040
            Width           =   1455
         End
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
            Height          =   1455
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   15975
         End
         Begin VB.ListBox List2 
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
            Height          =   1455
            Left            =   120
            TabIndex        =   2
            Top             =   5640
            Width           =   15975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"AjusteInventario.frx":2123
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
            Index           =   3
            Left            =   240
            TabIndex        =   18
            Top             =   5280
            Width           =   15735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"AjusteInventario.frx":21BB
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
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   15855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "UDM                        Disponible                                Real                                        Diferiencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   3600
            Width           =   12975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"AjusteInventario.frx":2250
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   2760
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
    Option Explicit
    
    '//RECORDSET
    Dim Rs      As New adodb.Recordset
    Dim Rs1     As New adodb.Recordset
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
        
        Frame1(1).Enabled = False
        Frame3(0).Enabled = True
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        List1.Clear
        List2.Clear
        
        With Rs2
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select distinct tipo from MTL_ON_HAND_QUANTITIES order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    frmAjusteInventario.Combo1.AddItem .Fields(0)
                    .MoveNext
                Wend
                
                .MoveFirst
                
                frmAjusteInventario.Combo1.Text = .Fields(0)
                
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
                If Mid(List1.Text, 1, 10) <> "" Then
                    For i = 0 To (7)
                        Text1(i) = ""
                    Next i
                    
                    Text1(0) = Trim(Mid(List1.Text, 1, 10))
                    Text1(4) = Trim(Mid(List1.Text, 68, 14))
                    Text1(7) = Trim(Mid(List1.Text, 94, 19))
                Else
                    MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
                End If
                    
            Case 1
                If Text1(0) <> "" Then
                    If Val(Text1(6)) <> 0 Then
                        c1 = Mid(Text1(0), 1, 10)
                        c2 = Mid(Text1(1), 1, 10)
                        c3 = Mid(Text1(2), 1, 44)
                        c4 = Mid(Text1(6), 1, 14)
                        c5 = Mid(Text1(3), 1, 10)
                        c6 = Mid(Text1(7), 1, 19)
                        
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
                        
                        For X = 0 To List1.ListCount - 1
                            If UCase(Trim(Mid(List2.List(X), 1, 10))) = UCase(Trim(Mid(Text1(0), 1, 10))) Then
                                If UCase(Trim(Mid(List2.List(X), 94, 19))) = UCase(Trim(Mid(Text1(7), 1, 19))) Then
                                    MsgBox UCase(Trim(Mid(List2.List(X), 94, 19)))
                                    MsgBox UCase(Trim(Mid(Text1(7), 1, 19)))
                                    MsgBox "El articulo ya esta en la lista", vbOKOnly, "Atención"
                                    
                                    For i = 0 To 6
                                        Text1(i) = ""
                                    Next i
                                    
                                    Exit Sub
                                End If
                            End If
                        Next
                        
                        List2.AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5 & " " & c6
                    End If
                    
                    For i = 0 To 7
                        Text1(i) = ""
                    Next i
                Else
                    MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
                End If
                    
            Case 2
                With List2
                    intX = .ListIndex
                    
                    .RemoveItem intX
                End With
                
            Case 3
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from MTL_ON_HAND_QUANTITIES where tipo = '" & frmAjusteInventario.Combo1.Text & "' order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
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
                            
                            List1.AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5 & " " & c6
                            
                            .MoveNext
                        Loop
                    End If
                End With
                
            Frame1(1).Enabled = True
            
            Frame3(0).Visible = False
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
                If Text1(0) = "" Then
                    For i = 1 To 6
                        Text1(i) = ""
                    Next i
                Else
                    Text1(1) = Get_ItemCod(Text1(0))
                    Text1(2) = Get_ItemDesc(Text1(0))
                    Text1(3) = Get_ItemUDM(Text1(0))
                End If
                    
            Case 5
                If Text1(5) = "" Then
                    Text1(6) = 0
                    Text1(5).BackColor = COLOR_NO_ENCONTRADO
                Else
                    Text1(6) = Replace(Format(Val(Text1(5)) - Val(Text1(4)), "0.00"), ",", ".")
                    Text1(5).BackColor = COLOR_NORMAL
                End If
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
            With Rs1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from MTL_MATERIAL_TRANSACTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Filter = ""
                .Requery
                
                If List2.ListCount > 0 Then
                    For i = 0 To List2.ListCount - 1
                        List2.ListIndex = i
                        List2.SetFocus
                        
                        v1 = Trim(Mid(List2.Text, 1, 10))
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
                        v10 = Trim(Mid(List2.Text, 94, 19))
                        
                        .AddNew
                            .Fields(1) = v1
                            .Fields(2) = v2
                            .Fields(3) = v3
                            .Fields(4) = v4
                            .Fields(5) = v5
                            .Fields(6) = v6
                            .Fields(7) = v7
                            .Fields(8) = v8
                            .Fields(9) = v9
                            If v10 <> "" Then
                                .Fields(10) = v10
                            End If
                        .Update
                        .Requery
                    Next i
                    
                    MsgBox "Ajuste Finalizado", vbOKOnly, "Información"
                Else
                    MsgBox "No se agregaron artículos para ajustar", vbCritical, "Advertencia"
                    
                    Exit Sub
                End If
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
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Rs2.State = 1 Then Rs2.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
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
