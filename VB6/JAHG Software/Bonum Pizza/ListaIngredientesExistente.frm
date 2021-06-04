VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaIngredientesExistente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Consultar listas de ingredientes existentes"
   ClientHeight    =   7260
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13905
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
   ScaleHeight     =   7260
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13455
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3375
            Left            =   240
            TabIndex        =   10
            Top             =   3120
            Width           =   13035
            _ExtentX        =   22992
            _ExtentY        =   5953
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            HeadLines       =   2
            RowHeight       =   28
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
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
            Height          =   420
            Index           =   0
            Left            =   1560
            TabIndex        =   9
            Top             =   240
            Width           =   11655
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
            Enabled         =   0   'False
            Height          =   420
            Index           =   1
            Left            =   1560
            TabIndex        =   8
            Top             =   720
            Width           =   11655
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   0
            Left            =   1560
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   1200
            Width           =   11655
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
            Index           =   2
            Left            =   1560
            TabIndex        =   3
            Top             =   1680
            Width           =   4455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "ListaIngredientesExistente.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2280
            Width           =   1455
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
            Left            =   -600
            TabIndex        =   11
            Top             =   240
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
            Top             =   1680
            Width           =   1815
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
            TabIndex        =   6
            Top             =   1200
            Width           =   1935
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
            TabIndex        =   5
            Top             =   720
            Width           =   2055
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
Attribute VB_Name = "frmListaIngredientesExistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset

Dim vItemMPId As Integer
' Constantes para indicar el color de fondo del combobox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const COLOR_NO_ENCONTRADO = &HC0C0FF ' color cuando no se encontró
Const COLOR_NORMAL = &HE0E0E0 ' color cuando hay coincidencia

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim i As Integer
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs
        
        If .State = 1 Then .Close
        .Open "Select t1.*,t2.descripcion + ' (' + t2.udm + ')' as nombre from ListasDeIngredientes t1, Items t2 where  t1.ItemPTId = t2.id and tipo = '" & StTipoItem & "' order by 3,6;", Cn, adOpenStatic, adLockOptimistic
        .Requery
        
        If .RecordCount <> 0 Then
            .MoveFirst
            .Filter = "ItemPTId = " & .Fields(0)
        End If
    
    End With
    
    With Rs3
        
        If .State = 1 Then .Close
        .Open "Select t1.*,t2.descripcion + ' (' + t2.udm + ')' as nombre from ListasDeIngredientes t1, Items t2 where  t1.ItemPTId = t2.id and tipo = '" & StTipoItem & "' order by 3,6;", Cn, adOpenStatic, adLockOptimistic
        .Requery
        
        If .RecordCount <> 0 Then
            .MoveFirst
            .Filter = "ItemPTId = " & .Fields(0)
        End If
    
    End With
    
    If Rs3.RecordCount > 0 Then
    
        With DataGrid1
            Set .DataSource = Rs3
            .Columns(0).Visible = False
            .Columns(1).Visible = False
            .Columns(2).Visible = False
            .Columns(3).Visible = False
            .Columns(7).Visible = False
            
            .Columns(4).Width = 2000
            .Columns(5).Width = 7000
            .Columns(6).Width = 3000
            
            .Columns(4).Caption = "Codigo"
            .Columns(5).Caption = "Descripcion"
            .Columns(6).Caption = "Cantidad"
            
            .Columns(4).Locked = True
            .Columns(5).Locked = True
        End With
        
        With Text1(1)
            Set .DataSource = Rs
            .DataField = "nombre"
        End With
        
        With Rs1
            
            If .State = 1 Then .Close
            .Open "Select t1.* from items t1 where t1.tipo = 'Ingredientes Generales' and not exists  (select 1 from ListasDeIngredientes where t1.id = ItemMPId and ItemPTId =" & Rs.Fields(0).Value & ") order by 3;", Cn, adOpenStatic, adLockOptimistic
            .Requery
            Combo1(0).Clear
            
            If .RecordCount <> 0 Then
                While Not .EOF
                    Combo1(0).AddItem .Fields(2) & " (" & .Fields(5) & ")"
                    .MoveNext
                Wend
            End If
        
        End With
        
        With Text1(2)
            .BackColor = &HC0C0FF
            .Text = ""
        End With
        
        Combo1(0).BackColor = &HC0C0FF
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Información"
        frmMenuInicial.Enabled = True
        Unload Me
        
    End If

End Sub

Private Sub Text1_Change(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            With Rs
                If Text1(0) = "" Then
                    .Filter = "ItemPTId = " & Rs.Fields(0)
                    .Requery
                Else
                    If .EOF = False Then
                        .Filter = "nombre like '*" & Text1(0) & "*'"
                        .Requery
                    Else
                        .Filter = ""
                        .Requery
                        .Filter = "ItemPTId = " & Rs.Fields(0)
                        .Requery
                    End If
                End If
            End With
            
            With Rs1
            
                If .State = 0 Then
                    .Open "Select t1.* from items t1 where t1.tipo = 'Ingredientes Generales' and not exists  (select 1 from ListasDeIngredientes where t1.id = ItemMPId and ItemPTId =" & Rs.Fields(0).Value & ") order by 3;", Cn, adOpenStatic, adLockOptimistic
                End If
                
                .Requery
                Combo1(0).Clear
                
                If .RecordCount <> 0 Then
                    While Not .EOF
                        Combo1(0).AddItem .Fields(2) & " (" & .Fields(5) & ")"
                        .MoveNext
                    Wend
                End If
            
            End With
            
        Case 1
            
            If Text1(1) <> "" Then
                
                With Rs3
                    .Filter = "nombre like '*" & Text1(1) & "*'"
                    .Requery
                End With
            
            Else
            
                With Rs
                    .Filter = ""
                    .Requery
                    .Filter = "ItemPTId = " & Rs.Fields(0)
                    .Requery
                End With
            
            End If
            
            With DataGrid1
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
                .Columns(3).Visible = False
                .Columns(7).Visible = False
                
                .Columns(4).Width = 2000
                .Columns(5).Width = 7000
                .Columns(6).Width = 3000
                
                .Columns(4).Caption = "Codigo"
                .Columns(5).Caption = "Descripcion"
                .Columns(6).Caption = "Cantidad"
                
                .Columns(4).Locked = True
                .Columns(5).Locked = True
            End With
        
        Case 2
            
            With Text1(2)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
    
    End Select

End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    On Error Resume Next
    
    Static cadena As String
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
                Else
                    
                    ' Backcolor normal cuando hay coincidencia
                    .BackColor = COLOR_NORMAL
                    .BackColor = &HE0E0E0
                    vItemMPId = Get_ItemId(.Text)
                    
                    With Rs1
                        .Filter = ""
                        .Filter = "Id = " & vItemMPId
                        .Requery
                        .MoveFirst
                    End With
                
                End If
            
            End With
    
    End Select

End Sub

Private Sub Combo1_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            With Combo1(0)
                
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    
                    .BackColor = &HE0E0E0
                    vItemMPId = Get_ItemId(.Text)
                    
                    With Rs1
                        .Filter = ""
                        .Filter = "Id = " & vItemMPId
                        .Requery
                        .MoveFirst
                    End With
                
                End If
            
            End With
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)
    
    On Error Resume Next
        
        Select Case Index
            
            Case 0
                
                If Combo1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" Then
                    
                    With Rs2
                        If .State = 1 Then .Close
                        .Open "Select * from ListasDeIngredientes", Cn, adOpenStatic, adLockOptimistic
                        .Requery
                        .AddNew
                            .Fields(0) = Rs.Fields(0).Value
                            .Fields(1) = Rs.Fields(1).Value
                            .Fields(2) = Rs.Fields(2).Value
                            .Fields(3) = Rs1.Fields(0).Value
                            .Fields(4) = Rs1.Fields(1).Value
                            .Fields(5) = Rs1.Fields(2).Value
                            .Fields(6) = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                        .Update
                        .Requery
                        .Close
                    End With
                    
                    Unload frmListaIngredientesExistente
                    Set frmListaIngredientesExistente = Nothing
                    frmListaIngredientesExistente.Show
                
                Else
                    MsgBox "Llenar todos los campos", vbCritical, "Error"
                End If
        
    End Select

End Sub

Private Sub Guardar_Click()
    
    On Error Resume Next
    
    With Rs3
        .Update
        .Requery
    End With
    
    With DataGrid1
        .Columns(0).Visible = False
        .Columns(1).Visible = False
        .Columns(2).Visible = False
        .Columns(3).Visible = False
        .Columns(7).Visible = False
                
        .Columns(4).Width = 2000
        .Columns(5).Width = 7000
        .Columns(6).Width = 3000
                
        .Columns(4).Caption = "Codigo"
        .Columns(5).Caption = "Descripcion"
        .Columns(6).Caption = "Cantidad"
                
        .Columns(4).Locked = True
        .Columns(5).Locked = True
    End With
    
    With Text1(1)
        Set .DataSource = Rs
        .DataField = "nombre"
    End With

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
    If Rs3.State = 1 Then Rs3.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Rs1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Cn = Nothing

End Sub



