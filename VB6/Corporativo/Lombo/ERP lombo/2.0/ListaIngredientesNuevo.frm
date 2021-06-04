VERSION 5.00
Begin VB.Form frmListaIngredientesNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear nueva lista de ingredientes"
   ClientHeight    =   7260
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   10410
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
   ScaleHeight     =   7260
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   7
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
            Height          =   3735
            Left            =   240
            TabIndex        =   5
            Top             =   2880
            Width           =   9375
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "ListaIngredientesNuevo.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   1800
            Picture         =   "ListaIngredientesNuevo.frx":0834
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1800
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
            TabIndex        =   2
            Top             =   1200
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
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   1560
            TabIndex        =   1
            Text            =   "Combo1"
            Top             =   720
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
            Left            =   360
            TabIndex        =   11
            Top             =   2520
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
            TabIndex        =   10
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ingrediente"
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
            Index           =   4
            Left            =   -480
            TabIndex        =   9
            Top             =   840
            Width           =   1935
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
            TabIndex        =   8
            Top             =   1320
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
Attribute VB_Name = "frmListaIngredientesNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim vItemPTId As Integer
Dim vItemMPId As Integer
Dim St As String
Dim v1 As Integer
Dim v2 As String
Dim v3 As String
Dim v4 As Integer
Dim v5 As String
Dim v6 As String
Dim v7 As String

' Constantes para indicar el color de fondo del combobox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const COLOR_NO_ENCONTRADO = &HC0C0FF ' color cuando no se encontró
Const COLOR_NORMAL = &HE0E0E0 ' color cuando hay coincidencia

Private Sub Form_Load()
    
    Dim i As Integer
    
    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs1
        
        If .State = 1 Then .Close
        
        .Open "Select * from items t1 where UDM <> 'Servicio' and not exists  (select 1 from ListasDeIngredientes where t1.id =ItemPTId) order by 3;", Cn, adOpenStatic, adLockOptimistic
        .Filter = ""
        .Requery
        
        If .RecordCount <> 0 Then
            
            Combo1(0).Clear
            
            While Not .EOF
                Combo1(0).AddItem .Fields(2) & " (" & .Fields(5) & ")"
                .MoveNext
            Wend
            
        End If
    
    End With
    
    If Rs1.RecordCount > 0 Then
        
        With Rs2
            
            If .State = 1 Then .Close
            .Open "Select t1.* from items t1 where UDM <> 'Servicio' order by 3;", Cn, adOpenStatic, adLockOptimistic
            .Filter = ""
            .Requery
                
            If .RecordCount <> 0 Then
                
                Combo1(1).Clear
                
                While Not .EOF
                    Combo1(1).AddItem .Fields(2) & " (" & .Fields(5) & ")"
                    .MoveNext
                Wend
            
            End If
        
        End With
        
        With Text1(0)
            .BackColor = &HC0C0FF
            .Text = ""
        End With
        
        For i = 0 To 1
            Combo1(i).BackColor = &HC0C0FF
        Next i
        
        List1.Clear
        
        Combo1(0).Text = ""
        Combo1(1).Text = ""
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Información"
        frmMenuInicial.Enabled = True
        Unload Me
        
    End If

End Sub

Private Sub Combo1_Click(Index As Integer)
    
    On Error Resume Next
    
    Dim cadena As String
    Dim i As Long
    
    Select Case Index
        
        Case 0
            
            With Combo1(0)
                
                If .Text = "" Then
                    
                    .BackColor = COLOR_NO_ENCONTRADO
                    
                    With Rs1
                        .Filter = ""
                        .Requery
                    End With
                
                Else
                    
                    .BackColor = COLOR_NORMAL
                    
                    With Combo1(0)
                        vItemPTId = Get_ItemId(.Text)
                    End With
                    
                    With Rs1
                        
                        .Filter = "Id = " & vItemPTId
                        .Requery
                        
                        v1 = .Fields(0).Value
                        v2 = .Fields(1).Value
                        v3 = .Fields(2).Value
                    
                    End With
                
                End If
            
            End With
        
        Case 1
            
            With Combo1(1)
                
                If .Text <> "" Then
                    
                    .BackColor = COLOR_NORMAL
                    
                    vItemMPId = Get_ItemId(Combo1(1).Text)
                    
                    With Rs2
                        .Filter = ""
                        .Filter = "Id = " & vItemMPId
                        .Requery
                    End With
                
                Else
                    .BackColor = COLOR_NO_ENCONTRADO
                End If
            
            End With
    
    End Select

End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    Dim cadena As String
    Dim i As Long
    
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
                    
                    With Rs1
                        .Filter = ""
                        .Requery
                    End With
                
                Else
                    
                    ' Backcolor normal cuando hay coincidencia
                    .BackColor = COLOR_NORMAL
                    
                    With Combo1(0)
                        vItemPTId = Get_ItemId(.Text)
                    End With
                    
                    With Rs1
                        
                        .Filter = "Id = " & vItemPTId
                        .Requery
                        
                        v1 = .Fields(0).Value
                        v2 = .Fields(1).Value
                        v3 = .Fields(2).Value
                    
                    End With
                
                End If
            
            End With
        
        Case 1
            
            With Combo1(1)
                
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
                    
                    vItemMPId = Get_ItemId(Combo1(1).Text)
                    
                    With Rs2
                        .Filter = ""
                        .Filter = "Id = " & vItemMPId
                        .Requery
                    End With
                
                End If
            
            End With
    
    End Select

End Sub

Private Sub Text1_Change(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            With Text1(0)
                
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            
            End With
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)
    
    Dim i As Integer
    Dim X As Integer
    
    On Error Resume Next
        
    Select Case Index
        
        Case 0
            
            If Combo1(0) <> "" And Combo1(1) <> "" And Text1(0) <> "" Then
                
                Dim viid As String
                Dim videscripcion As String
                Dim vicantidad As String
                Dim c1 As Integer
                Dim c2 As Integer
                Dim c3 As Integer
                
                viid = Rs2.Fields(0).Value
                videscripcion = Mid(Rs2.Fields(2).Value, 1, 43)
                vicantidad = Replace(Format(Val(Text1(0)), "0.00"), ",", ".")
                
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
                
                Text1(0).BackColor = &HC0C0FF
                
                With Combo1(1)
                    .Text = ""
                    .BackColor = &HC0C0FF
                    .SetFocus
                End With
            
            Else
                
                MsgBox "Llenar todos los campos", vbCritical, "Error"
                Text1(0).SetFocus
                
            End If
        
        Case 1
            
            Dim intX As Integer
            
            With List1
                intX = .ListIndex
                .RemoveItem (intX)
            End With
    
    End Select

End Sub

Private Sub Guardar_Click()
    
    Dim i As Integer
    
    On Error Resume Next
    
    With Rs
        
        If .State = 1 Then .Close
        
        .Open "Select * from ListasDeIngredientes;", Cn, adOpenStatic, adLockOptimistic
        .Filter = ""
        .Requery
        
        For i = 0 To List1.ListCount - 1
            
            List1.ListIndex = i
            List1.SetFocus
            
            v4 = Trim(Mid(List1.Text, 1, 10))
            v5 = Get_ItemCod(v4)
            v6 = Get_ItemDesc(v4)
            v7 = Replace(Trim(Mid(List1.Text, 56, 11)), ",", ".")
            
            .AddNew
                .Fields(0) = v1
                .Fields(1) = v2
                .Fields(2) = v3
                .Fields(3) = v4
                .Fields(4) = v5
                .Fields(5) = v6
                .Fields(6) = v7
            .Update
            .Requery
        
        Next i
    
    End With
    
    Unload frmListaIngredientesNuevo
    Set frmListaIngredientesNuevo = Nothing
    frmListaIngredientesNuevo.Show

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    frmMenuInicial.Enabled = True
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Rs1.State = 1 Then Rs1.Close
    If Rs2.State = 1 Then Rs2.Close
    
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Rs1 = Nothing
    Set Rs2 = Nothing
    
    Set Cn = Nothing

End Sub
