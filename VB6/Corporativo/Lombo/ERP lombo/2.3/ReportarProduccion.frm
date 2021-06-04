VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmReportarProduccion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Reportar Producción"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   13935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   13455
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   8760
            TabIndex        =   12
            Top             =   240
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1560
            TabIndex        =   1
            Top             =   720
            Width           =   11655
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "ReportarProduccion.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   1560
            TabIndex        =   2
            Top             =   1200
            Width           =   4455
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1560
            TabIndex        =   0
            Top             =   240
            Width           =   4455
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3375
            Left            =   240
            TabIndex        =   4
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
            ColumnCount     =   11
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
            BeginProperty Column02 
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
            BeginProperty Column03 
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
            BeginProperty Column04 
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
            BeginProperty Column05 
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
            BeginProperty Column06 
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
            BeginProperty Column07 
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
            BeginProperty Column08 
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
            BeginProperty Column09 
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
            BeginProperty Column10 
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
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
               EndProperty
               BeginProperty Column07 
               EndProperty
               BeginProperty Column08 
               EndProperty
               BeginProperty Column09 
               EndProperty
               BeginProperty Column10 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
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
            Left            =   6600
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Consumos de materiales"
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
            Left            =   600
            TabIndex        =   10
            Top             =   2640
            Width           =   3615
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
            TabIndex        =   9
            Top             =   720
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
            TabIndex        =   8
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Trabajo"
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
            TabIndex        =   7
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
    Option Explicit
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset  'folio
    Dim Rs1                 As New adodb.Recordset  'PT
    Dim Rs2                 As New adodb.Recordset  'PT
    Dim Rs3                 As New adodb.Recordset  'transacciones de inventario
    Dim Rs4                 As New adodb.Recordset  'trabajos de produccion
    Dim Rs5                 As New adodb.Recordset  'lote
    
    '//ARTICULOS
    Dim vlItemPTId          As Long
    
    '//OTROS
    Dim i                   As Long
    
    '//LOTE
    Dim ControlLote         As Long
    Dim InLoteExiste        As Long
    Dim CantidadRestante    As String
    Dim vLote               As String
    Dim vCantidadLote       As String
    Dim vCurrentLote        As String
    
    '//PRODUCCION
    Dim IdProduccion        As Long
    
    
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
            
            Text1(0).Text = "P-" & IdProduccion
        End With
        
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select distinct t2.descripcion + ' (' + t2.udm + ')' + ' (' + t2.codigo + ')' as nombre from BILL_OF_MATERIAL t1, MTL_SYSTEM_ITEMS t2 where t1.cantidad <> 0 and t1.ItemPTId = t2.id order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Filter = ""
            .Requery
            
            If .RecordCount <> 0 Then
                Combo1.Clear
                
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
        
        Text1(2).BackColor = COLOR_NO_ENCONTRADO
        
        Combo1.BackColor = COLOR_NO_ENCONTRADO
        
        With DataGrid1
            For i = 0 To 10
                .Columns(i).Visible = False
            Next i
            
            With .Columns(3)
                .Visible = True
                .Width = 6000
                .Locked = True
            End With
            
            With .Columns(6)
                .Visible = True
                .Width = 2000
                '.Text = Replace(Format(.Text, "0.00"), ",", ".")
            End With
            
            With .Columns(7)
                .Visible = True
                .Width = 2000
                .Locked = True
            End With
            
            With .Columns(10)
                .Visible = True
                .Width = 2000
            End With
        End With
        
        With Rs4
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from WIP_DISCONTINUE_JOBS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        vCurrentLote = "P" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdProduccion
        
        Text1(1) = vCurrentLote
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
    
    Private Sub Command1_Click(Index As Integer)
        On Error GoTo errHandler
        
        If Text1(2) = "" Or Val(Text1(2)) <= 0 Then
            MsgBox "Cantidad no válida", vbCritical, "Error"
                    
            Exit Sub
        End If
        
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
            
            Text1(0).Text = "P-" & IdProduccion
        End With
        
        vCurrentLote = "P" & Mid(Date, 9, 2) & Format(Date, "WW", vbMonday) & IdProduccion
        
        Text1(1) = vCurrentLote
                
        If Text1(2) <> "" And vlItemPTId <> 0 Then
            With Rs2
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from BILL_OF_MATERIAL where cantidad <> 0 and ItemPTId = " & vlItemPTId & ";", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                
                If .RecordCount <> 0 Then
                    While Not .EOF
                        'lote
                        ControlLote = Get_ItemLote(.Fields(3).Value)
                            
                        'si tiene control de lote
                        If ControlLote = 1 Then
                            CantidadRestante = Replace(Format(.Fields(6).Value * Val(Replace(Format(Val(Text1(2)), "0.00"), ",", ".")), "0.00"), ",", ".")
                            
                            'mientras no se complete la cantidad necesaria
                            While Val(CantidadRestante) > 0
                                'obtenemos lote mas antiguo y existencia de ese lote
                                vLote = ""
                                vLote = Get_LoteConsumo(.Fields(3).Value)
                                vCantidadLote = Get_LoteConsumoCantidad(.Fields(3).Value)
                                    
                                Rs3.AddNew
                                    Rs3.Fields(1) = .Fields(3).Value 'id
                                    Rs3.Fields(2) = .Fields(4).Value 'codigo
                                    Rs3.Fields(3) = .Fields(5).Value 'descripcion
                                    Rs3.Fields(4) = Date 'fecha
                                    Rs3.Fields(5) = "Consumo de Ingredientes" 'tipo de treansaccion
                                    Rs3.Fields(7) = Get_ItemUDM(.Fields(3).Value) 'udm
                                    Rs3.Fields(8) = Text1(0) 'folio
                                    Rs3.Fields(9) = "No" 'cancelado"
                                    
                                    'si existe algun lote
                                    If vLote <> "" Then
                                        If Val(vCantidadLote) > Val(CantidadRestante) Then
                                            Rs3.Fields(10) = vLote 'lote
                                            Rs3.Fields(6) = Val(CantidadRestante) * -1 'cantidad
                                            
                                            CantidadRestante = "0"
                                        Else
                                            Rs3.Fields(10) = vLote 'lote
                                            Rs3.Fields(6) = Replace(Format(Val(vCantidadLote) * -1, "0.00"), ",", ".") 'cantidad
                                            
                                            CantidadRestante = Replace(Format(Val(CantidadRestante) - Val(vCantidadLote), "0.00"), ",", ".")
                                        End If
                                    'si no existen lotes creamos uno
                                    Else
                                        InLoteExiste = Get_LoteExiste(vCurrentLote, .Fields(3).Value)
                                            
                                        If InLoteExiste = 0 Then
                                            If Rs5.State = 1 Then Rs5.Close
                                            Rs5.CursorLocation = adodb.CursorLocationEnum.adUseClient
                                            Rs5.Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                            Rs5.Requery
                                            Rs5.AddNew
                                                Rs5.Fields(1) = .Fields(3).Value 'idarticulo
                                                Rs5.Fields(2) = vCurrentLote 'lote
                                                Rs5.Fields(3) = "Produccion" 'tipo
                                            Rs5.Update
                                            Rs5.Requery
                                            Rs5.Close
                                        End If
                                            
                                        Rs3.Fields(10) = vCurrentLote 'lote
                                        Rs3.Fields(6) = Replace(Format(Val(CantidadRestante) * -1, "0.00"), ",", ".") 'cantidad
                                        
                                        CantidadRestante = "0"
                                    End If
                                Rs3.Update
                                Rs3.Requery
                            Wend
                        Else
                            Rs3.AddNew
                                Rs3.Fields(1) = .Fields(3).Value 'id
                                Rs3.Fields(2) = .Fields(4).Value 'codigo
                                Rs3.Fields(3) = .Fields(5).Value 'descripcion
                                Rs3.Fields(4) = Date 'fecha
                                Rs3.Fields(5) = "Consumo de Ingredientes" 'tipo de treansaccion
                                Rs3.Fields(6) = Replace(Format(.Fields(6).Value * Val(Replace(Format(Val(Text1(2)), "0.00"), ",", ".")) * -1, "0.00"), ",", ".") 'cantidad
                                Rs3.Fields(7) = Get_ItemUDM(.Fields(3).Value) 'udm
                                Rs3.Fields(8) = Text1(0) 'folio
                                Rs3.Fields(9) = "No" 'cancelado"
                            Rs3.Update
                            Rs3.Requery
                        End If
                        
                        .MoveNext
                    Wend
                    
                    ControlLote = Get_ItemLote(vlItemPTId)
                    'guardar lote
                    InLoteExiste = Get_LoteExiste(vCurrentLote, vlItemPTId)
                        
                    If ControlLote = 1 And InLoteExiste = 0 Then
                        With Rs5
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from MTL_LOT_NUMBERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                .Fields(1) = vlItemPTId 'idarticulo
                                .Fields(2) = vCurrentLote 'lote
                                .Fields(3) = "Produccion" 'tipo
                            .Update
                            .Requery
                            .Close
                        End With
                    End If
                End If
                
                Rs3.AddNew
                    Rs3.Fields(1) = vlItemPTId 'id
                    Rs3.Fields(2) = Get_ItemCod(vlItemPTId) 'codigo
                    Rs3.Fields(3) = Get_ItemDesc(vlItemPTId) 'descripcion
                    Rs3.Fields(4) = Date 'fecha
                    Rs3.Fields(5) = "Producción" 'tipo de transaccion
                    Rs3.Fields(6) = Replace(Format(Val(Text1(2)), "0.00"), ",", ".") 'cantidad
                    Rs3.Fields(7) = Get_ItemUDM(vlItemPTId) 'udm
                    Rs3.Fields(8) = Text1(0) 'folio
                    Rs3.Fields(9) = "No" 'cancelado"
                    
                    If ControlLote = 1 Then
                        Rs3.Fields(10) = vCurrentLote 'lote
                    End If
                        
                Rs3.Update
                Rs3.Requery
                
                .Close
            End With
            
            With Rs4
                .AddNew
                    .Fields(1) = Text1(0)
                .Update
                .Requery
            End With
            
            With Rs3
                .Filter = "Folio = '" & Text1(0) & "'"
                .Requery
            End With
            
            With DataGrid1
                Set .DataSource = Rs3
                
                For i = 0 To 10
                    .Columns(i).Visible = False
                Next i
                
                With .Columns(3)
                    .Visible = True
                    .Width = 6000
                    .Locked = True
                End With
                
                With .Columns(6)
                    .Visible = True
                    .Width = 2000
                End With
                
                With .Columns(7)
                    .Visible = True
                    .Width = 2000
                    .Locked = True
                End With
                
                With .Columns(10)
                    .Visible = True
                    .Width = 2000
                End With
            End With
    
            'Text1(2) = ""
            
            Combo1.Enabled = False
            
            Text1(2).Enabled = False
            
            Command1(0).Enabled = False
        Else
            MsgBox "Llenar todos los campos", vbCritical, "Error"
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
                    
        If vbq = vbYes Then
            With Rs3
                .Update
                .Requery
            End With
            
            MsgBox "Produccion guardada", vbOKOnly, "Terminado"
            
            Set DataGrid1.DataSource = Nothing
            
            Unload Me
            
            Set frmReportarProduccion = Nothing
            
            frmReportarProduccion.Show
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        'On Error GoTo errHandler
        
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmReportarProduccion:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        Set DataGrid1.DataSource = Nothing
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Rs2.State = 1 Then Rs2.Close
        If Rs3.State = 1 Then Rs3.Close
        If Rs4.State = 1 Then Rs4.Close
        If Rs5.State = 1 Then Rs5.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
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
