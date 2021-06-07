VERSION 5.00
Begin VB.Form frmListaIngredientesExistente 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Consultar listas de ingredientes existentes"
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
         Height          =   8655
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   16935
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "PRIMERO"
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
            Index           =   11
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   7920
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ANTERIOR"
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
            Index           =   12
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   7920
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "SIGUIENTE"
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
            Index           =   13
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   7920
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "ULTIMO"
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
            Index           =   14
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   7920
            Width           =   1575
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
            Left            =   2160
            TabIndex        =   7
            Top             =   6240
            Width           =   4455
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
            Index           =   5
            Left            =   2160
            TabIndex        =   9
            Top             =   7200
            Width           =   4455
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
            Left            =   2160
            TabIndex        =   8
            Top             =   6720
            Width           =   14655
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
            Height          =   3165
            Left            =   120
            TabIndex        =   6
            Top             =   2880
            Width           =   16695
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
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
            Height          =   420
            Index           =   1
            Left            =   15360
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1920
            Width           =   1455
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
            Index           =   0
            Left            =   2040
            TabIndex        =   1
            Top             =   120
            Visible         =   0   'False
            Width           =   13095
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
            Left            =   2040
            TabIndex        =   14
            Top             =   120
            Width           =   13095
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
            Left            =   2040
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   600
            Width           =   14775
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
            Index           =   2
            Left            =   2040
            TabIndex        =   4
            Top             =   1120
            Width           =   4455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION"
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
            Index           =   8
            Left            =   -480
            TabIndex        =   23
            Top             =   6720
            Width           =   2415
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
            Index           =   7
            Left            =   -240
            TabIndex        =   22
            Top             =   7200
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
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
            Left            =   -120
            TabIndex        =   21
            Top             =   6240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "INGREDIENTES"
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
            Left            =   240
            TabIndex        =   20
            Top             =   2520
            Width           =   12615
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
            Left            =   -240
            TabIndex        =   19
            Top             =   120
            Visible         =   0   'False
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
            Left            =   0
            TabIndex        =   18
            Top             =   1120
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INGREDIENTE"
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
            Left            =   -600
            TabIndex        =   17
            Top             =   600
            Width           =   2415
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
            Left            =   -240
            TabIndex        =   16
            Top             =   120
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
'***********************************************************************************
'Nombre:        frmListaIngredientesExistente
'Proposito:     Modificacion de listas
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Se agrego confirmacion de salida sin
'                                               guardar datos
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset
Dim RS1 As New adodb.Recordset
Dim Rs2 As New adodb.Recordset
Dim Rs3 As New adodb.Recordset
'//OTROS
Dim i As Long
Dim vItemMPId As Long
Dim vCategoria As String

Private Sub Form_Load()
    On Error GoTo errHandler
    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With Rs
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select t1.*,t2.descripcion + ' (' + t2.udm + ')' + ' (' + t2.codigo + ')' as nombre from BILL_OF_MATERIAL t1, MTL_SYSTEM_ITEMS t2 where  t1.ItemPTId = t2.id order by 4,7;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        If .RecordCount <> 0 Then
            .MoveFirst
            .Filter = "ItemPTId = " & .Fields(1)
        End If
    End With

    With List1
        .Clear
    End With

    With Rs3
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select t1.*,t2.descripcion + ' (' + t2.udm + ')' + ' (' + t2.codigo + ')' as nombre from BILL_OF_MATERIAL t1, MTL_SYSTEM_ITEMS t2 where  t1.ItemPTId = t2.id order by 3,6;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        If .RecordCount <> 0 Then
            .MoveFirst
            .Filter = "ItemPTId = " & .Fields(1)
        End If

        If .RecordCount > 0 Then
            .MoveFirst
            While Not Rs3.EOF
                List1.AddItem .Fields(6).Value & " (" & .Fields(0).Value & ")"
                .MoveNext
            Wend
            .MoveFirst
            With Text1(3)
                Set .DataSource = Rs3
                .DataField = "id"
            End With

            With Text1(4)
                Set .DataSource = Rs3
                .DataField = "ItemMPDescripcion"
            End With

            With Text1(5)
                Set .DataSource = Rs3
                .DataField = "Cantidad"
            End With

            With Text1(1)
                .Text = Rs.Fields("nombre")
            End With

            With RS1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select t1.* from MTL_SYSTEM_ITEMS t1 where t1.UDM <> 'Servicio' and t1.categoria = 'Inventario' and id <> " & Rs.Fields(0).Value & " and not exists (select 1 from BILL_OF_MATERIAL where t1.id = itemMPId and ItemPTId =" & Rs.Fields(0).Value & ") order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                With Combo1(0)
                    .Clear
                End With

                If .RecordCount <> 0 Then
                    While Not .EOF
                        Combo1(0).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                        .MoveNext
                    Wend
                End If
            End With

            With Text1(2)
                .BackColor = COLOR_NO_ENCONTRADO
                .Text = ""
            End With

            With Combo1(0)
                .BackColor = COLOR_NO_ENCONTRADO
            End With
        Else
            MsgBox "No hay registros existentes", vbOKOnly, "Información"
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Form_Load" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Text1_Change(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 1
        With Combo1(0)
            .Clear
        End With

        With Text1(1)
            If .Text = "" Then
                With Rs
                    .Filter = ""
                    .Requery
                    .Filter = "ItemPTId = " & .Fields(0)
                    .Requery
                End With
            Else
                With Rs
                    .Filter = "nombre like '*" & Text1(1) & "*'"
                    .Requery
                End With

                With Rs3
                    .Filter = "nombre like '*" & Text1(1) & "*'"
                    .Requery
                End With
            End If
        End With

        With RS1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select t1.* from MTL_SYSTEM_ITEMS t1 where T1.UDM <> 'Servicio' and t1.Categoria = 'Inventario' and id <> " & Rs.Fields(0).Value & " and not exists (select 1 from BILL_OF_MATERIAL where t1.id = itemMPId and ItemPTId =" & Rs.Fields(0).Value & ") order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount <> 0 Then
                While Not .EOF
                    Combo1(0).AddItem .Fields(2) & " (" & .Fields(9) & ")" & " (" & .Fields(1) & ")"
                    .MoveNext
                Wend
            End If
        End With

        With List1
            .Clear
        End With

        With Rs3
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    List1.AddItem .Fields(6).Value & " (" & .Fields(0).Value & ")"
                    .MoveNext
                Wend
                .MoveFirst
            End If
        End With
    Case 2
        With Text1(2)
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Text1_Change" & vbTab & err.Number & vbTab & err.Description
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
                vItemMPId = Get_ItemId(.Text)
                With RS1
                    .Filter = ""
                    .Filter = "Id = " & vItemMPId
                    .Requery
                    .MoveFirst
                End With
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Combo1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Combo1(0)
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
                vItemMPId = Get_ItemId(.Text)
                With RS1
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select t1.* from MTL_SYSTEM_ITEMS t1 where t1.UDM <> 'Servicio' and t1.categoria = 'Inventario' and id <> " & Rs.Fields(0).Value & " and not exists (select 1 from BILL_OF_MATERIAL where t1.id = itemMPId and ItemPTId =" & Rs.Fields(0).Value & ") order by 3;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Filter = "Id = " & vItemMPId
                    .Requery
                    .MoveFirst
                End With
            End If
        End With
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
    Case 0
        With Text1(2)
            If .Text = "" Or Val(.Text) <= 0 Then
                MsgBox "Cantidad no válida", vbCritical, "Error"
                Exit Sub
            End If
            If Combo1(0) <> "" And Text1(1) <> "" And .Text <> "" Then
                With Rs2
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from BILL_OF_MATERIAL", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .AddNew
                    With .Fields(1)
                        .Value = Rs.Fields(1).Value                                             'id pt
                    End With

                    With .Fields(2)
                        .Value = Rs.Fields(2).Value                                             'codigo pt
                    End With

                    With .Fields(3)
                        .Value = Rs.Fields(3).Value                                             'desc pt
                    End With

                    With .Fields(4)
                        .Value = RS1.Fields(0).Value                                            'id mp
                    End With

                    With .Fields(5)
                        .Value = RS1.Fields(1).Value                                            'codigo mp
                    End With

                    With .Fields(6)
                        .Value = RS1.Fields(2).Value                                            'desc mp
                    End With

                    With .Fields(7)
                        .Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")               'cantidad
                    End With

                    With .Fields(8)
                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                    End With

                    With .Fields(9)
                        .Value = StUsuario                                                      'usuario
                    End With

                    With .Fields(10)
                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                    End With

                    With .Fields(11)
                        .Value = StUsuario                                                      'usuario
                    End With
                    .Update
                    .Requery
                    .Close
                End With
                Unload frmListaIngredientesExistente
                Set frmListaIngredientesExistente = Nothing

                With frmListaIngredientesExistente
                    .Show
                End With
            Else
                MsgBox "Llenar todos los campos", vbCritical, "Error"
            End If
        End With
    Case 1
        With frmBuscadorListaIngredientes
            .Show 1
        End With
    Case 11
        With List1
            .ListIndex = 0
        End With
    Case 12
        With List1
            .ListIndex = List1.ListIndex - 1
        End With
    Case 13
        With List1
            .ListIndex = List1.ListIndex + 1
        End With
    Case 14
        With List1
            .ListIndex = List1.ListCount - 1
        End With
    End Select
    Exit Sub
errHandler:
    If err.Number = 380 Then
        err.Clear
        Exit Sub
    End If
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Command1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub List1_Click()
    On Error GoTo errHandler
    Dim Str As String
    Dim ArrStr() As String
    Dim FilterId As Integer

    With List1
        If .Text = "" Then
            MsgBox "Seleccione algún ingrediente", vbOKOnly, "Información"
        Else
            Str = .Text
            ArrStr() = Split(Str, "(")
            FilterId = Replace(ArrStr(1), ")", "")
            With Rs3
                .Filter = "id = " & FilterId
                .Requery
                .MoveFirst
            End With
        End If
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:List1_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Rs3
            With .Fields("last_updated_by")
                .Value = StUsuario
            End With

            With .Fields("last_update_date")
                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
            End With
            .Update
            .Requery
        End With

        With Text1(1)
            Set .DataSource = Rs

            .DataField = "nombre"
        End With
        Unload frmListaIngredientesExistente
        Set frmListaIngredientesExistente = Nothing

        With frmListaIngredientesExistente
            .Show
        End With
    End If
    Exit Sub
errHandler:
    If err.Number = 3219 Then
        With Rs3
            With .Fields("last_updated_by")
                .Value = StUsuario
            End With

            With .Fields("last_update_date")
                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
            End With
            .Update
            .Requery
        End With
        err.Clear
        With Text1(1)
            Set .DataSource = Rs
            .DataField = "nombre"
        End With
        Unload frmListaIngredientesExistente
        Set frmListaIngredientesExistente = Nothing

        With frmListaIngredientesExistente
            .Show
        End With

        Exit Sub
    End If
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Salir_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Rs3
            With .Fields("last_updated_by")
                .Value = StUsuario
            End With

            With .Fields("last_update_date")
                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
            End With
            .Update
            .Requery
        End With
    End If
    Unload Me
    Exit Sub
errHandler:
    If err.Number = 3219 Then
        With Rs3
            With .Fields("last_updated_by")
                .Value = StUsuario
            End With

            With .Fields("last_update_date")
                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
            End With
            .Update
            .Requery
        End With
        err.Clear
        Unload Me
        Exit Sub
    End If
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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

    With Rs3
        If .State = 1 Then .Close
    End With

    With Cn
        If .State = 1 Then .Close
    End With

    Set Rs = Nothing
    Set RS1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Cn = Nothing
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmListaIngredientesExistente:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
