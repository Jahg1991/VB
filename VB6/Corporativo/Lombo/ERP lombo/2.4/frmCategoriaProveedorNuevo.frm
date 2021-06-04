VERSION 5.00
Begin VB.Form frmCategoriaProveedorNuevo 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Nueva Categoria de Articulo"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   160.073
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   307.182
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   7650
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   16695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1920
      MaxLength       =   255
      TabIndex        =   1
      Top             =   360
      Width           =   15135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
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
Attribute VB_Name = "frmCategoriaProveedorNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmCategoriaProveedorNuevo
'Proposito:     Registro de Categorias para Proveedores
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
    '//OTROS
    Dim i       As Long
    Dim In1     As Long
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        With Text1
            .BackColor = COLOR_NO_ENCONTRADO
        End With
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from HZ_PARTY_CATEGORIES order by 2;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    List1.AddItem .Fields(1).Value
                    .MoveNext
                Wend
            End If
        End With
        
        With Cn
            .Close
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCategoriaProveedorNuevo:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Text1_Change()
        On Error GoTo errHandler
        With Text1
            If .Text = "" Then
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCategoriaProveedorNuevo:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
        If vbq = vbYes Then
            With Text1
                If .Text <> "" Then
                    With Cn
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        If .State = 0 Then .Open (StConnection)
                    End With
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select count(*) as existe from HZ_PARTY_CATEGORIES where categoria like '" & Text1 & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        In1 = .Fields(0).Value
                        .Close
                    End With
                    
                    If In1 = 0 Then
                        With Rs
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from HZ_PARTY_CATEGORIES;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                With .Fields(1)
                                    .Value = Text1
                                End With
                                
                                With .Fields(2)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
                                End With
                                
                                With .Fields(3)
                                    .Value = StUsuario
                                End With
                                
                                With .Fields(4)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
                                End With
                                
                                With .Fields(5)
                                    .Value = StUsuario
                                End With
                            .Update
                            .Requery
                            .Close
                        End With
                        Unload frmCategoriaProveedorNuevo
                        Set frmCategoriaProveedorNuevo = Nothing
                        
                        With frmCategoriaProveedorNuevo
                            .Show
                        End With
                    Else
                        MsgBox "La categoria ya existe", vbCritical, "Error"
                        With Text1
                            .SetFocus
                        End With
                    End If
                    
                    With Cn
                        If .State = 1 Then .Close
                    End With
                Else
                    MsgBox "Llenar todos los campos", vbCritical, "Error"
                    .SetFocus
                End If
            End With
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCategoriaProveedorNuevo:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCategoriaProveedorNuevo:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        With Rs
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCategoriaProveedorNuevo:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
