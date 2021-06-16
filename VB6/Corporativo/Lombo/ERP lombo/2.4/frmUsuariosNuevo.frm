VERSION 5.00
Begin VB.Form frmUsuariosNuevo 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Nuevo Usuario"
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
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   20
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   8520
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "REPORTES"
      ForeColor       =   &H00C0C000&
      Height          =   5280
      Index           =   0
      Left            =   10800
      TabIndex        =   37
      Top             =   1440
      Width           =   6255
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   12
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   13
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   14
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   15
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2280
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   16
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   17
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3480
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   18
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   4080
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   19
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   4680
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CATALOGOS"
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   14
         Left            =   -2160
         TabIndex        =   45
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA DE MAT."
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   15
         Left            =   -2160
         TabIndex        =   44
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCCION"
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   16
         Left            =   0
         TabIndex        =   43
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   18
         Left            =   -360
         TabIndex        =   42
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDOS DE V."
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   19
         Left            =   0
         TabIndex        =   41
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COMPRAS"
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   20
         Left            =   480
         TabIndex        =   40
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTARIOS"
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   21
         Left            =   480
         TabIndex        =   39
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CORTE DE CAJA"
         ForeColor       =   &H00C0C000&
         Height          =   375
         Index           =   22
         Left            =   600
         TabIndex        =   38
         Top             =   4680
         Width           =   2295
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   11
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   7920
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   10
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   7320
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   9
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   6720
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   8
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   6120
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   7
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5520
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   6
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4920
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   5
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   4
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3720
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   3
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3120
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   2
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2520
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   1
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   0
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5520
      MaxLength       =   255
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   840
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   5520
      MaxLength       =   255
      TabIndex        =   14
      Top             =   360
      Width           =   11535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESARROLLADOR"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   17
      Left            =   240
      TabIndex        =   46
      Top             =   8520
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA PREDETERMINADA"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   13
      Left            =   240
      TabIndex        =   13
      Top             =   7920
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PERMITIR CAMBIO DE CAJA"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   12
      Left            =   240
      TabIndex        =   12
      Top             =   7320
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A CORTE DE CAJA"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   11
      Left            =   240
      TabIndex        =   11
      Top             =   6720
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A INVENTARIOS"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   6120
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A AJUSTES DE INVENTARIO"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   5520
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A COMPRAS"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A PEDIDOS DE VENTA"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A VENTAS"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A PRODUCCION"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A LISTAS DE MATERIALES"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A CATALOGOS"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESO A ARCHIVO"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
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
Attribute VB_Name = "frmUsuariosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmUsusariosNuevo
'Proposito:     Registro de usuarios
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Se agrego confirmacion de salida sin
'                                               guardar datos
'
'1.2        16/06/2021     Alfredo Hernandez    Se agrego la responsabilidad de
'                                               Desarrollador
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset
'//OTROS
Dim i As Long
Dim In1 As Long

Private Sub Form_Load()
    On Error GoTo errHandler
    For i = 0 To 1
        With Text1(i)
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    For i = 0 To 10
        With Combo1(i)
            .AddItem "Si"
            .AddItem "No"
            .Text = "No"
        End With
    Next i

    With Combo1(11)
        For i = 1 To 10
            .AddItem "Caja " & i
        Next i
        .Text = "Caja 1"
    End With

    For i = 12 To 20
        With Combo1(i)
            .AddItem "Si"
            .AddItem "No"
            .Text = "No"
        End With
    Next i
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmUsuariosNuevo:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
                .BackColor = COLOR_NO_ENCONTRADO
            Else
                .BackColor = COLOR_NORMAL
            End If
        End With
    Case 1
        With Text1(1)
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmUsuariosNuevo:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Text1(0)
            If .Text <> "" And Text1(1) <> "" Then
                With Cn
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    If .State = 0 Then .Open (StConnection)
                End With

                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from FND_USERS where nombre like '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    In1 = .Fields(0).Value
                    .Close
                    If In1 = 0 Then
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from FND_USERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                        With .Fields(1)
                            .Value = Text1(0)                                                       'usuario
                        End With

                        With .Fields(2)
                            .Value = Text1(1)                                                       'pass
                        End With

                        With .Fields(3)
                            .Value = Combo1(0)                                                      'Archivo
                        End With

                        With .Fields(4)
                            .Value = Combo1(1)                                                      'Catalogos
                        End With

                        With .Fields(5)
                            .Value = Combo1(2)                                                      'Listas
                        End With

                        With .Fields(6)
                            .Value = Combo1(3)                                                      'Produccion
                        End With

                        With .Fields(7)
                            .Value = Combo1(4)                                                      'Ventas
                        End With

                        With .Fields(8)
                            .Value = Combo1(5)                                                      'Pedidos
                        End With

                        With .Fields(9)
                            .Value = Combo1(6)                                                      'Compras
                        End With

                        With .Fields(10)
                            .Value = Combo1(7)                                                      'Ajustes
                        End With

                        With .Fields(11)
                            .Value = Combo1(8)                                                      'Inventario
                        End With

                        With .Fields(12)
                            .Value = Combo1(9)                                                      'Corte
                        End With

                        With .Fields(13)
                            .Value = Combo1(10)                                                     'Caja
                        End With

                        With .Fields(14)
                            .Value = Combo1(11)                                                     'Caja Predeterminada
                        End With

                        With .Fields(15)
                            .Value = Combo1(12)                                                     'RCacalogos
                        End With

                        With .Fields(16)
                            .Value = Combo1(13)                                                     'RListas
                        End With

                        With .Fields(17)
                            .Value = Combo1(14)                                                     'RProduccion
                        End With

                        With .Fields(18)
                            .Value = Combo1(15)                                                     'RPedidos
                        End With

                        With .Fields(19)
                            .Value = Combo1(16)                                                     'RVentas
                        End With

                        With .Fields(20)
                            .Value = Combo1(17)                                                     'RCompras
                        End With

                        With .Fields(21)
                            .Value = Combo1(18)                                                     'RInventario
                        End With

                        With .Fields(22)
                            .Value = Combo1(19)                                                     'RCorte
                        End With
                        
                        With .Fields(23)
                            .Value = Combo1(20)                                                     'Desarrollador
                        End With

                        With .Fields(24)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                        End With

                        With .Fields(25)
                            .Value = StUsuario                                                      'usuario
                        End With

                        With .Fields(26)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                        End With

                        With .Fields(27)
                            .Value = StUsuario                                                      'usuario
                        End With
                        .Update
                        .Requery
                        .Close
                        Unload frmUsuariosNuevo
                        Set frmUsuariosNuevo = Nothing

                        With frmUsuariosNuevo
                            .Show
                        End With
                    Else
                        MsgBox "El usuario ya existe", vbCritical, "Error"
                        With Text1(0)
                            .SetFocus
                        End With
                    End If
                End With

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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmUsuariosNuevo:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Salir_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Text1(0)
            If .Text <> "" And Text1(1) <> "" Then
                With Cn
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    If .State = 0 Then .Open (StConnection)
                End With

                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from FND_USERS where nombre like '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    In1 = .Fields(0).Value
                    .Close
                    If In1 = 0 Then
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select * from FND_USERS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        .AddNew
                        With .Fields(1)
                            .Value = Text1(0)                                                       'usuario
                        End With

                        With .Fields(2)
                            .Value = Text1(1)                                                       'pass
                        End With

                        With .Fields(3)
                            .Value = Combo1(0)                                                      'Archivo
                        End With

                        With .Fields(4)
                            .Value = Combo1(1)                                                      'Catalogos
                        End With

                        With .Fields(5)
                            .Value = Combo1(2)                                                      'Listas
                        End With

                        With .Fields(6)
                            .Value = Combo1(3)                                                      'Produccion
                        End With

                        With .Fields(7)
                            .Value = Combo1(4)                                                      'Ventas
                        End With

                        With .Fields(8)
                            .Value = Combo1(5)                                                      'Pedidos
                        End With

                        With .Fields(9)
                            .Value = Combo1(6)                                                      'Compras
                        End With

                        With .Fields(10)
                            .Value = Combo1(7)                                                      'Ajustes
                        End With

                        With .Fields(11)
                            .Value = Combo1(8)                                                      'Inventario
                        End With

                        With .Fields(12)
                            .Value = Combo1(9)                                                      'Corte
                        End With

                        With .Fields(13)
                            .Value = Combo1(10)                                                     'Caja
                        End With

                        With .Fields(14)
                            .Value = Combo1(11)                                                     'Caja Predeterminada
                        End With

                        With .Fields(15)
                            .Value = Combo1(12)                                                     'RCacalogos
                        End With

                        With .Fields(16)
                            .Value = Combo1(13)                                                     'RListas
                        End With

                        With .Fields(17)
                            .Value = Combo1(14)                                                     'RProduccion
                        End With

                        With .Fields(18)
                            .Value = Combo1(15)                                                     'RPedidos
                        End With

                        With .Fields(19)
                            .Value = Combo1(16)                                                     'RVentas
                        End With

                        With .Fields(20)
                            .Value = Combo1(17)                                                     'RCompras
                        End With

                        With .Fields(21)
                            .Value = Combo1(18)                                                     'RInventario
                        End With

                        With .Fields(22)
                            .Value = Combo1(19)                                                     'RCorte
                        End With

                        With .Fields(23)
                            .Value = Combo1(20)                                                     'Desarrollador
                        End With

                        With .Fields(24)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                        End With

                        With .Fields(25)
                            .Value = StUsuario                                                      'usuario
                        End With

                        With .Fields(26)
                            .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                        End With

                        With .Fields(27)
                            .Value = StUsuario                                                      'usuario
                        End With
                        .Update
                        .Requery
                        .Close
                        Unload frmUsuariosNuevo
                        Set frmUsuariosNuevo = Nothing
                    Else
                        MsgBox "El usuario ya existe", vbCritical, "Error"
                        With Text1(0)
                            .SetFocus
                        End With
                    End If
                End With

                With Cn
                    If .State = 1 Then .Close
                End With
            Else
                MsgBox "Llenar todos los campos", vbCritical, "Error"
                .SetFocus
            End If
        End With
    Else
        Unload Me
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmUsuariosNuevo:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmUsuariosNuevo:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
