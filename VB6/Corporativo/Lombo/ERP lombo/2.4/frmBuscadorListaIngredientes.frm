VERSION 5.00
Begin VB.Form frmBuscadorListaIngredientes 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Lista de materiales"
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
         Height          =   9015
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   0
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
            Top             =   8160
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
            Top             =   840
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
            Top             =   240
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
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmBuscadorListaIngredientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmBuscadorListaIngredientes
'Proposito:     Buscar listas
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
    
    Dim Rs As New adodb.Recordset
    
    Sub Form_Load()
        On Error GoTo errHandler
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select DISTINCT t2.descripcion + ' (' + t2.udm + ')' + ' (' + t2.codigo + ')' as nombre from BILL_OF_MATERIAL t1, MTL_SYSTEM_ITEMS t2 where  t1.ItemPTId = t2.id order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount <> 0 Then
                .MoveFirst
                With List1
                    .Clear
                End With
                
                While Not .EOF
                    List1.AddItem .Fields(0).Value
                    .MoveNext
                Wend
            Else
                Unload Me
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorListaIngredientes:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
                        List1.AddItem .Fields(0).Value
                        .MoveNext
                    Wend
                End If
            Else
                .Filter = "nombre like '*" & Text1 & "*'"
                .Requery
                If .RecordCount <> 0 Then
                    .MoveFirst
                    While Not .EOF
                        List1.AddItem .Fields(0).Value
                        .MoveNext
                    Wend
                End If
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorListaIngredientes:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        With List1
            If .Text = "" Then
                MsgBox "Seleccione algun articulo", vbOKOnly, "Información"
            Else
                With frmListaIngredientesExistente
                    With .Text1(1)
                        .Text = List1.Text
                    End With
                End With
                Unload Me
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorListaIngredientes:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub List1_DblClick()
        On Error GoTo errHandler
        With List1
            If .Text = "" Then
                MsgBox "Seleccione algun articulo", vbOKOnly, "Información"
            Else
                With frmListaIngredientesExistente
                    With .Text1(1)
                        .Text = List1.Text
                    End With
                End With
                Unload Me
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorListaIngredientes:List1_DblClick" & vbTab & err.Number & vbTab & err.Description
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
        
        Set Rs = Nothing
        Set frmBuscadorListaIngredientes = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorListaIngredientes:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
