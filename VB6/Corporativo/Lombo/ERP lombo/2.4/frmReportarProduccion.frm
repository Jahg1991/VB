VERSION 5.00
Begin VB.Form frmReportarProduccion 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Reportar Producción"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
         Width           =   16800
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
            Index           =   7
            Left            =   13320
            TabIndex        =   17
            Top             =   7200
            Width           =   3500
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   7920
            Width           =   1575
         End
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
            TabIndex        =   18
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
            Left            =   8040
            TabIndex        =   16
            Top             =   7200
            Width           =   3500
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
            Height          =   2475
            Left            =   480
            TabIndex        =   12
            Top             =   3960
            Width           =   16335
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
            Left            =   8040
            TabIndex        =   14
            Top             =   6720
            Width           =   8775
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
            Left            =   1800
            TabIndex        =   15
            Top             =   7200
            Width           =   3500
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
            Left            =   1800
            TabIndex        =   13
            Top             =   6720
            Width           =   3500
         End
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
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2280
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
            Left            =   1680
            TabIndex        =   10
            Top             =   720
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
            Left            =   1680
            TabIndex        =   1
            Top             =   1200
            Width           =   15135
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
            Left            =   1680
            TabIndex        =   2
            Top             =   1720
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
            Index           =   0
            Left            =   1680
            TabIndex        =   0
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "UDM"
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
            Index           =   10
            Left            =   5640
            TabIndex        =   27
            Top             =   7200
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LOTE"
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
            Index           =   9
            Left            =   10920
            TabIndex        =   26
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
            Left            =   -480
            TabIndex        =   25
            Top             =   6720
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
            Index           =   7
            Left            =   -600
            TabIndex        =   24
            Top             =   7200
            Width           =   2175
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
            Left            =   5400
            TabIndex        =   23
            Top             =   6720
            Width           =   2415
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
            Index           =   4
            Left            =   480
            TabIndex        =   22
            Top             =   3480
            Width           =   7005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LOTE"
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
            Left            =   -600
            TabIndex        =   9
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CONSUMO DE MATERIALES"
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
            TabIndex        =   8
            Top             =   3000
            Width           =   3615
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
            TabIndex        =   7
            Top             =   1200
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
            TabIndex        =   6
            Top             =   1720
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TRABAJO"
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
            Left            =   -600
            TabIndex        =   5
            Top             =   240
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
Attribute VB_Name = "frmReportarProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmReportarProduccion
'Proposito:     Registro de Produccion, entrada de PT y consumo de MP registrados
'               en la lista
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Se agrego confirmacion de salida sin
'                                               guardar datos
'
'1.2        14/05/2021     Alfredo Hernandez    Se agrego funcion botones primero,
'                                               anterior, siguiente y ultimo
'
'1.3        14/05/2021     Alfredo Hernandez    Se agrego validacion servicio e
'                                               inventario para consumos
'
'1.4        14/05/2021     Alfredo Hernandez    Se agrego validacion servicio e
'                                               inventario para produccion
'
'1.5        14/05/2021     Alfredo Hernandez    Se agrego cambio de color a Text1(2)
'
'1.6        18/05/2021     Alfredo Hernandez    Se borran registros si el usuario
'                                               decide no guardar
'
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset  'folio
    Dim RS1                 As New adodb.Recordset  'PT
    Dim Rs2                 As New adodb.Recordset  'PT
    Dim Rs3                 As New adodb.Recordset  'transacciones de inventario
    Dim Rs4                 As New adodb.Recordset  'trabajos de produccion
    Dim Rs5                 As New adodb.Recordset  'lote
    '//ARTICULOS
    Dim vlItemPTId          As Long
    '//OTROS
    Dim i                   As Long
    '//LOTE
    Dim ControlLote         As Boolean
    Dim InLoteExiste        As Long
    Dim CantidadRestante    As String
    Dim vLote               As String
    Dim vCantidadLote       As String
    Dim vCurrentLote        As String
    '//PRODUCCION
    Dim IdProduccion        As Long
    
    Dim Str                 As String
    Dim ArrStr()            As String
    Dim FilterId            As Integer
    '//ELIMINAR
    Dim sql                 As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from WIP_TRANSACTION_ID", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            .MoveFirst
            If IsNull(Rs!IdProduccion) = False Then
                IdProduccion = Rs!IdProduccion
            Else
                IdProduccion = 1
            End If
            
            With Text1(0)
                .Text = "P-" & IdProduccion
            End With
        End With
        
        With RS1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select distinct t2.descripcion + ' (' + t2.udm + ')' + ' (' + t2.codigo + ')' as nombre from BILL_OF_MATERIAL t1, MTL_SYSTEM_ITEMS t2 where t1.cantidad <> 0 and t1.ItemPTId = t2.id order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            If .RecordCount <> 0 Then
                With Combo1
                    .Clear
                End With
                
                While Not .EOF
                    Combo1.AddItem .Fields(0)
                    .MoveNext
                Wend
            End If
        End With
        With Rs3
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from MTL_MATERIAL_TRANSACTIONS where TipoTransaccion = 'Consumo de Ingredientes';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Text1(2)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
        
        With Combo1
            .BackColor = COLOR_NO_ENCONTRADO
        End With
        
        With List1
            .Clear
        End With
        
        With Rs4
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from WIP_DISCONTINUE_JOBS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        vCurrentLote = "P" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdProduccion
        With Text1(1)
            .Text = vCurrentLote
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
        On Error GoTo errHandler
        Static cadena As String
        
        With Combo1
            ' si pesionamos las teclas de las flechas sale de la rutina
            If KeyCode = vbKeyUp Then Exit Sub
            
            If KeyCode = vbKeyDown Then Exit Sub
            
            If KeyCode = vbKeyLeft Then Exit Sub
            
            If KeyCode = vbKeyRight Then Exit Sub
            
            ' verifica que no se presionó la tecla backspace
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
                vlItemPTId = 0
            Else
                ' Backcolor normal cuando hay coincidencia
                .BackColor = COLOR_NORMAL
                vlItemPTId = Get_ItemId(.Text)
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Combo1_KeyUp" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_Click()
        On Error GoTo errHandler
        With Combo1
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
                vlItemPTId = 0
            Else
                .BackColor = COLOR_NORMAL
                vlItemPTId = Get_ItemId(.Text)
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Text1_Change(Index As Integer)
        On Error GoTo errHandler
        Select Case Index
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Text1_Change" & vbTab & err.Number & vbTab & err.Description
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
                End With
                'Actualizar folio
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select * from WIP_TRANSACTION_ID", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    .MoveFirst
                    If IsNull(Rs!IdProduccion) = False Then
                        IdProduccion = Rs!IdProduccion
                    Else
                        IdProduccion = 1
                    End If
                    
                    With Text1(0)
                        .Text = "P-" & IdProduccion
                    End With
                End With
                vCurrentLote = "P" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdProduccion
                With Text1(1)
                    .Text = vCurrentLote
                End With
                
                With Text1(2)
                    If .Text <> "" And vlItemPTId <> 0 Then
                        With Rs2
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from BILL_OF_MATERIAL where cantidad <> 0 and ItemPTId = " & vlItemPTId & ";", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            If .RecordCount <> 0 Then
                                While Not .EOF
                                    If Get_ItemUDM(.Fields(4).Value) <> "Servicio" And Get_ItemCategoria(.Fields(4).Value) = "Inventario" Then
                                        If PcInventarios = False Then
                                            If Val(Get_CantidadItem(.Fields(4).Value)) < Val(Replace(Format(.Fields(7).Value * Val(Replace(Format(Val(Text1(2)), "0.00"), ",", ".")), "0.00"), ",", ".")) Then
                                                MsgBox "Existencia insuficiente, no se realizará el consumo del ingrediente " & Get_ItemDesc(.Fields(4).Value), vbCritical, "Advertencia"
                                                GoTo Siguiente
                                            End If
                                        End If
                                        'lote
                                        ControlLote = Get_ItemLote(.Fields(4).Value)
                                        'si tiene control de lote
                                        If ControlLote = True Then
                                            CantidadRestante = Replace(Format(.Fields(7).Value * Val(Replace(Format(Val(Text1(2)), "0.00"), ",", ".")), "0.00"), ",", ".")
                                            'mientras no se complete la cantidad necesaria
                                            While Val(CantidadRestante) > 0
                                                'obtenemos lote mas antiguo y existencia de ese lote
                                                vLote = ""
                                                vLote = Get_LoteConsumo(.Fields(4).Value)
                                                vCantidadLote = Get_LoteConsumoCantidad(.Fields(4).Value)
                                                With Rs3
                                                    .AddNew
                                                        With .Fields(1)
                                                            .Value = Rs2.Fields(4).Value                                                                            'id
                                                        End With
                                                        
                                                        With .Fields(2)
                                                            .Value = Rs2.Fields(5).Value                                                                            'codigo
                                                        End With
                                                        
                                                        With .Fields(3)
                                                            .Value = Rs2.Fields(6).Value                                                                            'descripcion
                                                        End With
                                                        
                                                        With .Fields(4)
                                                            .Value = Date                                                                                           'fecha
                                                        End With
                                                        
                                                        With .Fields(5)
                                                            .Value = "Consumo de Ingredientes"                                                                      'tipo de treansaccion
                                                        End With
                                                        
                                                        With .Fields(7)
                                                            .Value = Get_ItemUDM(Rs2.Fields(4).Value)                                                               'udm
                                                        End With
                                                        
                                                        With .Fields(8)
                                                            .Value = Text1(0)                                                                                       'folio
                                                        End With
                                                        
                                                        With .Fields(9)
                                                            .Value = "No"                                                                                           'cancelado"
                                                        End With
                                                        
                                                        'si existe algun lote
                                                        If vLote <> "" Then
                                                            If Val(vCantidadLote) > Val(CantidadRestante) Then
                                                                With .Fields(10)
                                                                    .Value = vLote                                                                                  'lote
                                                                End With
                                                                
                                                                With .Fields(6)
                                                                    .Value = Val(CantidadRestante) * -1                                                             'cantidad
                                                                End With
                                                                CantidadRestante = "0"
                                                            Else
                                                                With .Fields(10)
                                                                    .Value = vLote                                                                                  'lote
                                                                End With
                                                                
                                                                With .Fields(6)
                                                                    .Value = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".")                             'cantidad
                                                                End With
                                                                CantidadRestante = Replace(Format(Val(CantidadRestante) - Val(vCantidadLote), "0.00"), ",", ".")
                                                            End If
                                                        'si no existen lotes creamos uno
                                                        Else
                                                            InLoteExiste = Get_LoteExiste(vCurrentLote, Rs2.Fields(4).Value)
                                                            If InLoteExiste = 0 Then
                                                                With Rs5
                                                                    If .State = 1 Then .Close
                                                                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                                                    .Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                                                    .Requery
                                                                    .AddNew
                                                                        With .Fields(1)
                                                                            .Value = Rs2.Fields(4).Value                                            'idarticulo
                                                                        End With
                                                                        
                                                                        With .Fields(2)
                                                                            .Value = vCurrentLote                                                   'lote
                                                                        End With
                                                                        
                                                                        With .Fields(3)
                                                                            .Value = "Produccion"                                                   'tipo
                                                                        End With
                                                                        
                                                                        With .Fields("Creation_date")
                                                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                                                        End With
                                                                        
                                                                        With .Fields("Created_by")
                                                                            .Value = StUsuario                                                      'usuario
                                                                        End With
                                                                        
                                                                        With .Fields("Last_update_date")
                                                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                                                        End With
                                                                        
                                                                        With .Fields("last_updated_by")
                                                                            .Value = StUsuario                                                      'usuario
                                                                        End With
                                                                    .Update
                                                                    .Requery
                                                                    .Close
                                                                End With
                                                            End If
                                                            
                                                            With .Fields(10)
                                                                .Value = vCurrentLote                                                                                   'lote
                                                            End With
                                                            
                                                            With .Fields(6)
                                                                .Value = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".")                                  'cantidad
                                                            End With
                                                            CantidadRestante = "0"
                                                        End If
                                                        
                                                        With .Fields("Creation_date")
                                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")                                        'creacion
                                                        End With
                                                        
                                                        With .Fields("Created_by")
                                                            .Value = StUsuario                                                                                          'usuario
                                                        End With
                                                        
                                                        With .Fields("Last_update_date")
                                                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")                                        'modificacion
                                                        End With
                                                        
                                                        With .Fields("last_updated_by")
                                                            .Value = StUsuario                                                                                          'usuario
                                                        End With
                                                    .Update
                                                    .Requery
                                                End With
                                            Wend
                                        Else
                                            With Rs3
                                                .AddNew
                                                    With .Fields(1)
                                                        .Value = Rs2.Fields(4).Value                                                                                                    'id
                                                    End With
                                                    
                                                    With .Fields(2)
                                                        .Value = Rs2.Fields(5).Value                                                                                                    'codigo
                                                    End With
                                                    
                                                    With .Fields(3)
                                                        .Value = Rs2.Fields(6).Value                                                                                                    'descripcion
                                                    End With
                                                    
                                                    With .Fields(4)
                                                        .Value = Date                                                                                                                   'fecha
                                                    End With
                                                    
                                                    With .Fields(5)
                                                        .Value = "Consumo de Ingredientes"                                                                                              'tipo de treansaccion
                                                    End With
                                                    
                                                    With .Fields(6)
                                                        .Value = Replace(Format(Rs2.Fields(7).Value * Val(Replace(Format(Val(Text1(2)), "0.00"), ",", ".")) * -1, "0.00"), ",", ".")    'cantidad
                                                    End With
                                                    
                                                    With .Fields(7)
                                                        .Value = Get_ItemUDM(Rs2.Fields(4).Value)                                                                                       'udm
                                                    End With
                                                    
                                                    With .Fields(8)
                                                        .Value = Text1(0)                                                                                                               'folio
                                                    End With
                                                    
                                                    With .Fields(9)
                                                        .Value = "No"                                                                                                                   'cancelado"
                                                    End With
                                                    
                                                    With .Fields("Creation_date")
                                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")                                                            'creacion
                                                    End With
                                                    
                                                    With .Fields("Created_by")
                                                        .Value = StUsuario                                                                                                              'usuario
                                                    End With
                                                    
                                                    With .Fields("Last_update_date")
                                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")                                                            'modificacion
                                                    End With
                                                    
                                                    With .Fields("last_updated_by")
                                                        .Value = StUsuario                                                                                                              'usuario
                                                    End With
                                                .Update
                                                .Requery
                                            End With
                                        End If
                                    End If
Siguiente:
                                    .MoveNext
                                Wend
                                If Get_ItemUDM(vlItemPTId) = "Servicio" Or Get_ItemCategoria(vlItemPTId) <> "Inventario" Then
                                    GoTo trabajo
                                End If
                                ControlLote = Get_ItemLote(vlItemPTId)
                                'guardar lote
                                InLoteExiste = Get_LoteExiste(vCurrentLote, vlItemPTId)
                                If ControlLote = True And InLoteExiste = 0 Then
                                    With Rs5
                                        If .State = 1 Then .Close
                                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                        .Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                        .Requery
                                        .AddNew
                                            With .Fields(1)
                                                .Value = vlItemPTId                                                     'idarticulo
                                            End With
                                            
                                            With .Fields(2)
                                                .Value = vCurrentLote                                                   'lote
                                            End With
                                            
                                            With .Fields(3)
                                                .Value = "Produccion"                                                   'tipo
                                            End With
                                            
                                            With .Fields("Creation_date")
                                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                            End With
                                            
                                            With .Fields("Created_by")
                                                .Value = StUsuario                                                      'usuario
                                            End With
                                            
                                            With .Fields("Last_update_date")
                                                .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                            End With
                                            
                                            With .Fields("last_updated_by")
                                                .Value = StUsuario                                                      'usuario
                                            End With
                                        .Update
                                        .Requery
                                        .Close
                                    End With
                                End If
                            End If
                            With Rs3
                                .AddNew
                                    With .Fields(1)
                                        .Value = vlItemPTId                                                     'id
                                    End With
                                    
                                    With .Fields(2)
                                        .Value = Get_ItemCod(vlItemPTId)                                        'codigo
                                    End With
                                    
                                    With .Fields(3)
                                        .Value = Get_ItemDesc(vlItemPTId)                                       'descripcion
                                    End With
                                    
                                    With .Fields(4)
                                        .Value = Date                                                           'fecha
                                    End With
                                    
                                    With .Fields(5)
                                        .Value = "Producción"                                                   'tipo de transaccion
                                    End With
                                    
                                    With .Fields(6)
                                        .Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")               'cantidad
                                    End With
                                    
                                    With .Fields(7)
                                        .Value = Get_ItemUDM(vlItemPTId)                                        'udm
                                    End With
                                    
                                    With .Fields(8)
                                        .Value = Text1(0)                                                       'folio
                                    End With
                                    
                                    With .Fields(9)
                                        .Value = "No"                                                           'cancelado"
                                    End With
                                    
                                    If ControlLote = True Then
                                        With .Fields(10)
                                            .Value = vCurrentLote                                               'lote
                                        End With
                                    End If
                                    
                                    With .Fields("Creation_date")
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                    End With
                                    
                                    With .Fields("Created_by")
                                        .Value = StUsuario                                                      'usuario
                                    End With
                                    
                                    With .Fields("Last_update_date")
                                        .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                    End With
                                    
                                    With .Fields("last_updated_by")
                                        .Value = StUsuario                                                      'usuario
                                    End With
                                .Update
                                .Requery
                            End With
                            .Close
                        End With
trabajo:
                        With Rs4
                            .AddNew
                                With .Fields(1)
                                    .Value = Text1(0)                                                       'trabajo
                                End With
                                
                                With .Fields(2)
                                    .Value = vlItemPTId                                                     'id
                                End With
                                    
                                With .Fields(3)
                                    .Value = Get_ItemCod(vlItemPTId)                                        'codigo
                                End With
                                    
                                With .Fields(4)
                                    .Value = Get_ItemDesc(vlItemPTId)                                       'descripcion
                                End With
                                
                                With .Fields("Creation_date")
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                End With
                                
                                With .Fields("Created_by")
                                    .Value = StUsuario                                                      'usuario
                                End With
                                
                                With .Fields("Last_update_date")
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                End With
                                
                                With .Fields("last_updated_by")
                                    .Value = StUsuario                                                      'usuario
                                End With
                            .Update
                            .Requery
                        End With
                        
                        With Rs3
                            .Filter = "Folio = '" & Text1(0) & "'"
                            .Requery
                            With List1
                                .Clear
                            End With
                            
                            If .RecordCount <> 0 Then
                                .MoveFirst
                                While Not .EOF
                                    List1.AddItem .Fields(3) & " (" & .Fields(0) & ")"
                                    .MoveNext
                                Wend
                                .MoveFirst
                            End If
                        End With
                        
                        With Text1(3)
                            Set .DataSource = Rs3
                            .DataField = "Id"
                        End With
                        
                        With Text1(4)
                            Set .DataSource = Rs3
                            .DataField = "ItemDescricion"
                        End With
                        
                        With Text1(5)
                            Set .DataSource = Rs3
                            .DataField = "Cantidad"
                        End With
                        
                        With Text1(6)
                            Set .DataSource = Rs3
                            .DataField = "UDM"
                        End With
                        
                        With Text1(7)
                            Set .DataSource = Rs3
                            .DataField = "lote"
                        End With
                        
                        With Combo1
                            .Enabled = False
                        End With
                        .Enabled = False
                        With Command1(0)
                            .Enabled = False
                        End With
                    Else
                        MsgBox "Llenar todos los campos", vbCritical, "Error"
                    End If
                End With
            Case 11
                With List1
                    .ListIndex = 0
                End With
            Case 12
                With List1
                    .ListIndex = .ListIndex - 1
                End With
            Case 13
                With List1
                    .ListIndex = .ListIndex + 1
                End With
            Case 14
                With List1
                    .ListIndex = .ListCount - 1
                End With
        End Select
    Exit Sub
errHandler:
        If err.Number = 380 Then
            With List1
                .ListIndex = 0
            End With
            
            Exit Sub
        End If
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub List1_Click()
        On Error GoTo errHandler
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:List1_Click" & vbTab & err.Number & vbTab & err.Description
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
            MsgBox "Produccion guardada", vbOKOnly, "Terminado"
            Unload Me
            Set frmReportarProduccion = Nothing
            
            With frmReportarProduccion
                .Show
            End With
        Else
            sql = "DELETE FROM MTL_MATERIAL_TRANSACTIONS WHERE FOLIO LIKE '" & Text1(0) & "'"
            With Cn
                .Execute sql
            End With
            sql = "DELETE FROM WIP_DISCONTINUE_JOBS WHERE TRABAJO LIKE '" & Text1(0) & "'"
            With Cn
                .Execute sql
            End With
            sql = "DELETE FROM MTL_LOT_NUMBERS WHERE LOTE LIKE '" & vCurrentLote & "'"
            With Cn
                .Execute sql
            End With
        End If
    Exit Sub
errHandler:
        If err.Number = 3021 Then
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
            MsgBox "Produccion guardada", vbOKOnly, "Terminado"
            Unload Me
            Set frmReportarProduccion = Nothing
            
            With frmReportarProduccion
                .Show
            End With
            
            Exit Sub
        End If
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
            MsgBox "Produccion guardada", vbOKOnly, "Terminado"
        Else
            sql = "DELETE FROM MTL_MATERIAL_TRANSACTIONS WHERE FOLIO LIKE '" & Text1(0) & "'"
            With Cn
                .Execute sql
            End With
            sql = "DELETE FROM WIP_DISCONTINUE_JOBS WHERE TRABAJO LIKE '" & Text1(0) & "'"
            With Cn
                .Execute sql
            End With
            sql = "DELETE FROM MTL_LOT_NUMBERS WHERE LOTE LIKE '" & vCurrentLote & "'"
            With Cn
                .Execute sql
            End With
        End If
        Unload Me
    Exit Sub
errHandler:
        If err.Number = 3021 Then
            err.Clear
            MsgBox "Produccion guardada", vbOKOnly, "Terminado"
            Unload Me
            Set frmReportarProduccion = Nothing
            
            With frmReportarProduccion
                .Show
            End With
            Unload Me
        End If
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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
        
        With Rs4
            If .State = 1 Then .Close
        End With
        
        With Rs5
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set RS1 = Nothing
        Set Rs2 = Nothing
        Set Rs3 = Nothing
        Set Rs4 = Nothing
        Set Rs5 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
