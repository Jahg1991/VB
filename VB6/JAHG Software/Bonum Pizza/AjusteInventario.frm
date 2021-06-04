VERSION 5.00
Begin VB.Form frmAjusteInventario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ajuste de Inventario"
   ClientHeight    =   7665
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13965
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
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   7215
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13455
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   6
            Left            =   9720
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            Picture         =   "AjusteInventario.frx":0000
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
            Picture         =   "AjusteInventario.frx":08D1
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
            Picture         =   "AjusteInventario.frx":1105
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
            Width           =   13095
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
            Width           =   13095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"AjusteInventario.frx":1914
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
            TabIndex        =   17
            Top             =   5280
            Width           =   12855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Id                Código        Descripción                                                                       Disponible  UDM"
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
            TabIndex        =   16
            Top             =   120
            Width           =   12855
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
            Caption         =   "Id                     Código                    Descripción"
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
            Width           =   12975
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
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim i As Integer
Dim c1 As String
Dim c2 As String
Dim c3 As String
Dim c4 As String
Dim c5 As String
Dim nc As Integer
Dim intX As Integer
Dim v1 As Integer
Dim v2 As String
Dim v3 As String
Dim v4 As String
Dim v5 As String
Dim v6 As String
Dim v7 As String
Dim v8 As String
Dim v9 As String

Private Sub Form_Load()

    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    List1.Clear
    List2.Clear
    
    With Rs
        
        If .State = 1 Then .Close
        
        If StTipoItem = "Barra" Then
            .Open "Select * from stock where tipo = 'Barra' order by 3;", Cn, adOpenStatic, adLockOptimistic
        End If
        
        If StTipoItem = "Cocina" Then
            .Open "Select * from stock where tipo = 'Cocina' order by 3;", Cn, adOpenStatic, adLockOptimistic
        End If
        
        If StTipoItem = "Otros" Then
            .Open "Select * from stock where (tipo = 'Otros' or tipo = 'Ingredientes Generales') order by 3;", Cn, adOpenStatic, adLockOptimistic
        End If
        
        .Requery
    
        If .RecordCount > 0 Then
    
            Do Until .EOF
        
                c1 = Mid(.Fields(1).Value, 1, 10)
                c2 = Mid(.Fields(2).Value, 1, 10)
                c3 = Mid(.Fields(3).Value, 1, 44)
                c4 = Replace(Format(Mid(.Fields(4).Value, 1, 14), "0.00"), ",", ".")
                c5 = Mid(.Fields(5).Value, 1, 10)
                
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
                
                
                List1.AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5
    
                .MoveNext
            
            Loop
        
        Else
    
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
            frmMenuInicial.Enabled = True
            Unload Me
        
        End If
        
    End With
    
End Sub

Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            If Mid(List1.Text, 1, 10) <> "" Then
            
                For i = 0 To 6
                    Text1(i) = ""
                Next i
                
                Text1(0) = Trim(Mid(List1.Text, 1, 10))
                Text1(4) = Trim(Mid(List1.Text, 68, 14))
            
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
                    
                    For X = 0 To List1.ListCount - 1
                    
                        If UCase(Trim(Mid(List2.List(X), 1, 10))) = UCase(Trim(Mid(Text1(0), 1, 10))) Then
                            
                            MsgBox "El articulo ya esta en la lista", vbOKOnly, "Atención"
                            
                            For i = 0 To 6
                                Text1(i) = ""
                            Next i
                
                            Exit Sub
                        
                        End If
                        
                    Next
                    
                    List2.AddItem c1 & " " & c2 & " " & c3 & " " & c4 & " " & c5
                
                End If
                
                For i = 0 To 6
                    Text1(i) = ""
                Next i
            
            Else
            
                MsgBox "Seleccionar un elemento en la lista", vbCritical, "Advertencia"
            
            End If
            
        Case 2
        
            With List2
                intX = .ListIndex
                .RemoveItem (intX)
            End With
            
    End Select
    
End Sub

Private Sub Text1_Change(Index As Integer)

    On Error Resume Next
    
    Dim i As Integer
    
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
                Text1(5).BackColor = &HC0C0FF
                
            Else
                
                Text1(6) = Replace(Format(Val(Text1(5)) - Val(Text1(4)), "0.00"), ",", ".")
                Text1(5).BackColor = &HE0E0E0
            
            End If
    
    End Select

End Sub

Private Sub Guardar_Click()

    On Error Resume Next
    
        With Rs1
            
            If .State = 1 Then .Close
            
            .Open "Select * from TransaccionesDeInventario;", Cn, adOpenStatic, adLockOptimistic
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
                    v8 = "Ajuste " & StTipoItem
                    v9 = "No"
                        
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
                    .Update
                    .Requery
                
                Next i
                
                MsgBox "Ajuste Finalizado", vbOKOnly, "Información"
            
            Else
            
                MsgBox "No se agregaron artículos para ajustar", vbCritical, "Advertencia"
                
                Exit Sub
                
            End If
        
        End With
    
        frmMenuInicial.Enabled = True
        Unload Me
    
End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    frmMenuInicial.Enabled = True
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Sub


