VERSION 5.00
Begin VB.Form frmClientesNuevo 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Añadir Clientes"
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
      ForeColor       =   &H00404000&
      Height          =   8895
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17220
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   8655
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16935
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   14
            Left            =   3120
            TabIndex        =   20
            Top             =   4920
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   13
            Left            =   3120
            TabIndex        =   19
            Top             =   4440
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   12
            Left            =   3120
            TabIndex        =   18
            Top             =   3960
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   11
            Left            =   3120
            TabIndex        =   17
            Top             =   3480
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   10
            Left            =   3120
            TabIndex        =   16
            Top             =   3000
            Width           =   2415
         End
         Begin VB.ComboBox Combo2 
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
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   6880
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   9
            Left            =   3120
            TabIndex        =   27
            Top             =   7920
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   3120
            TabIndex        =   26
            Top             =   7420
            Width           =   2415
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
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   6360
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   22
            Top             =   5880
            Width           =   13695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   21
            Top             =   5400
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   15
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   14
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   13
            Top             =   1560
            Width           =   13695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   12
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   11
            Top             =   600
            Width           =   13695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
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
            Left            =   3120
            TabIndex        =   10
            Top             =   120
            Width           =   13695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO 6"
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
            Index           =   16
            Left            =   600
            TabIndex        =   35
            Top             =   4920
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO 5"
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
            Index           =   15
            Left            =   600
            TabIndex        =   34
            Top             =   4440
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO 4"
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
            Index           =   14
            Left            =   360
            TabIndex        =   33
            Top             =   3960
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO 3"
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
            Index           =   13
            Left            =   840
            TabIndex        =   32
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO 2"
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
            Index           =   12
            Left            =   600
            TabIndex        =   31
            Top             =   3000
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MAYORISTA"
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
            Index           =   11
            Left            =   120
            TabIndex        =   30
            Top             =   6880
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CREDITO (DIAS)"
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
            Left            =   120
            TabIndex        =   29
            Top             =   7920
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CREDITO ($)"
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
            Left            =   600
            TabIndex        =   28
            Top             =   7420
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "LISTA DE PRECIOS"
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
            Left            =   120
            TabIndex        =   23
            Top             =   6360
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "REFERENCIAS"
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
            Left            =   1080
            TabIndex        =   9
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NO. TARJETA"
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
            Left            =   600
            TabIndex        =   8
            Top             =   5400
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO 1"
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
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CODIGO POSTAL"
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
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COLONIA"
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
            Left            =   1440
            TabIndex        =   5
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NUMERO"
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
            Left            =   1440
            TabIndex        =   4
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CALLE"
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
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NOMBRE"
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
            Left            =   1440
            TabIndex        =   2
            Top             =   120
            Width           =   1455
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
Attribute VB_Name = "frmClientesNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmClientesNuevo
'Proposito:     Registro de clientes
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
    Dim Rs      As New adodb.Recordset
    '//OTROS
    Dim i       As Long
    Dim In1     As Long
    Dim sql     As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        For i = 0 To (7)
            With Text1(i)
                .BackColor = COLOR_NO_ENCONTRADO
            End With
        Next i
        
        For i = 10 To (14)
            With Text1(i)
                .BackColor = COLOR_NO_ENCONTRADO
            End With
        Next i
        
        For i = 1 To (5)
            With Combo1
                .AddItem i
            End With
        Next i
        
        With Combo1
            .Text = "1"
        End With
        
        With Combo2
            .AddItem "Si"
            .AddItem "No"
            .Text = "No"
        End With
        
        With Text1(8)
            .Text = "0"
        End With
        
        With Text1(9)
            .Text = "0"
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesNuevo:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
            Case 2
                With Text1(2)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 3
                With Text1(3)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 4
                With Text1(4)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 5
                With Text1(5)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 6
                With Text1(6)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 7
                With Text1(7)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 8
                With Text1(8)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 9
                With Text1(9)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 10
                With Text1(10)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 11
                With Text1(11)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 12
                With Text1(12)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 13
                With Text1(13)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            Case 14
                With Text1(14)
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesNuevo:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
        If vbq = vbYes Then
            If Text1(0) <> "" Then
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(cliente,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    In1 = .Fields(0).Value
                    .Close
                End With
                
                If In1 = 0 Then
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select count(*) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(proveedor,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        In1 = .Fields(0).Value
                        .Close
                    End With
                    
                    If In1 = 0 Then
                        With Rs
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from HZ_PARTY;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                With .Fields(1)
                                    .Value = Text1(0)                                                       'nombre
                                End With
                                
                                With .Fields(2)
                                    .Value = Text1(1)                                                       'calle
                                End With
                                
                                With .Fields(3)
                                    .Value = Text1(2)                                                       'numero
                                End With
                                
                                With .Fields(4)
                                    .Value = Text1(3)                                                       'colonia
                                End With
                                
                                With .Fields(5)
                                    .Value = Text1(4)                                                       'cp
                                End With
                                
                                With .Fields(6)
                                    .Value = Text1(5)                                                       'tel
                                End With
                                
                                With .Fields(7)
                                    .Value = Text1(10)                                                      'tel2
                                End With
                                
                                With .Fields(8)
                                    .Value = Text1(11)                                                      'tel3
                                End With
                                
                                With .Fields(9)
                                    .Value = Text1(12)                                                      'tel4
                                End With
                                
                                With .Fields(10)
                                    .Value = Text1(13)                                                      'tel5
                                End With
                                
                                With .Fields(11)
                                    .Value = Text1(14)                                                      'tel6
                                End With
                                
                                With .Fields(13)
                                    .Value = Text1(6)                                                       'monedero
                                End With
                                
                                With .Fields(14)
                                    .Value = Text1(7)                                                       'referencias
                                End With
                                
                                With .Fields(15)
                                    .Value = "Cliente"                                                      'tipo
                                End With
                                
                                With .Fields(16)
                                    .Value = Combo1                                                         'lista
                                End With
                                
                                With .Fields(17)
                                    .Value = Text1(8)                                                       'credito
                                End With
                                
                                With .Fields(18)
                                    .Value = Text1(9)                                                       'dias
                                End With
                                
                                With .Fields(19)
                                    .Value = Combo2                                                         'mayorista
                                End With
                                
                                With .Fields(21)
                                    .Value = "Si"                                                           'cliente
                                End With
                                
                                With .Fields(23)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                End With
                                
                                With .Fields(24)
                                    .Value = StUsuario                                                      'usuario
                                End With
                                
                                With .Fields(25)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                End With
                                
                                With .Fields(26)
                                    .Value = StUsuario                                                      'usuario
                                End With
                            .Update
                            .Requery
                            .Close
                        End With
                        
                        If InTipoAltaClienteProveedor = 1 Then
                            With frmVentas
                                .Enabled = True
                            End With
                            Unload Me
                            Set frmClientesNuevo = Nothing
                            
                            Exit Sub
                        Else
                            Unload frmClientesNuevo
                            Set frmClientesNuevo = Nothing
                            
                            With frmClientesNuevo
                                .Show
                            End With
                            
                            Exit Sub
                        End If
                    Else
                        vbq = MsgBox("El nombre ya está registrado como proveedor, ¿Desea convertirlo también en cliente?", vbQuestion + vbYesNo, "Información")
                        If vbq = vbYes Then
                            With Rs
                                If .State = 1 Then .Close
                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                .Open "Select min(id) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(proveedor,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                .Requery
                                In1 = .Fields(0).Value
                                .Close
                            End With
                            sql = "update HZ_PARTY set cliente = 'Si', [Lista de Precios] = " & Combo1 & ",Mayorista ='" & Combo2 & "',last_updated_by ='" & StUsuario & "', last_update_date = '" & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS") & "' where id = " & In1
                            With Cn
                                .Execute sql
                            End With
                            
                            If InTipoAltaClienteProveedor = 1 Then
                                With frmVentas
                                    .Enabled = True
                                End With
                                Unload Me
                                Set frmClientesNuevo = Nothing
                                
                                Exit Sub
                            Else
                                Unload frmClientesNuevo
                                Set frmClientesNuevo = Nothing
                                
                                With frmClientesNuevo
                                    .Show
                                End With
                                
                                Exit Sub
                            End If
                        Else
                            With Text1(0)
                                .SetFocus
                            End With
                            
                            Exit Sub
                        End If
                    End If
                Else
                    MsgBox "El cliente ya esta registrado", vbCritical, "Error"
                    With Text1(0)
                        .SetFocus
                    End With
                    
                    Exit Sub
                End If
                
                With Cn
                    If .State = 1 Then .Close
                End With
            Else
                MsgBox "El nombre es obligatorio", vbCritical, "Error"
                With Text1(0)
                    .SetFocus
                End With
                
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesNuevo:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
        If vbq = vbYes Then
            If Text1(0) <> "" Then
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(cliente,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    In1 = .Fields(0).Value
                    .Close
                End With
                
                If In1 = 0 Then
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select count(*) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(proveedor,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        In1 = .Fields(0).Value
                        .Close
                    End With
                    
                    If In1 = 0 Then
                        With Rs
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from HZ_PARTY;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                With .Fields(1)
                                    .Value = Text1(0)                                                       'nombre
                                End With
                                
                                With .Fields(2)
                                    .Value = Text1(1)                                                       'calle
                                End With
                                
                                With .Fields(3)
                                    .Value = Text1(2)                                                       'numero
                                End With
                                
                                With .Fields(4)
                                    .Value = Text1(3)                                                       'colonia
                                End With
                                
                                With .Fields(5)
                                    .Value = Text1(4)                                                       'cp
                                End With
                                
                                With .Fields(6)
                                    .Value = Text1(5)                                                       'tel
                                End With
                                
                                With .Fields(7)
                                    .Value = Text1(10)                                                      'tel2
                                End With
                                
                                With .Fields(8)
                                    .Value = Text1(11)                                                      'tel3
                                End With
                                
                                With .Fields(9)
                                    .Value = Text1(12)                                                      'tel4
                                End With
                                
                                With .Fields(10)
                                    .Value = Text1(13)                                                      'tel5
                                End With
                                
                                With .Fields(11)
                                    .Value = Text1(14)                                                      'tel6
                                End With
                                
                                With .Fields(13)
                                    .Value = Text1(6)                                                       'monedero
                                End With
                                
                                With .Fields(14)
                                    .Value = Text1(7)                                                       'referencias
                                End With
                                
                                With .Fields(15)
                                    .Value = "Cliente"                                                      'tipo
                                End With
                                
                                With .Fields(16)
                                    .Value = Combo1                                                         'lista
                                End With
                                
                                With .Fields(17)
                                    .Value = Text1(8)                                                       'credito
                                End With
                                
                                With .Fields(18)
                                    .Value = Text1(9)                                                       'dias
                                End With
                                
                                With .Fields(19)
                                    .Value = Combo2                                                         'mayorista
                                End With
                                
                                With .Fields(21)
                                    .Value = "Si"                                                           'cliente
                                End With
                                
                                With .Fields(23)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'creacion
                                End With
                                
                                With .Fields(24)
                                    .Value = StUsuario                                                      'usuario
                                End With
                                
                                With .Fields(25)
                                    .Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")    'modificacion
                                End With
                                
                                With .Fields(26)
                                    .Value = StUsuario                                                      'usuario
                                End With
                            .Update
                            .Requery
                            .Close
                        End With
                        
                        If InTipoAltaClienteProveedor = 1 Then
                            With frmVentas
                                .Enabled = True
                            End With
                            Unload Me
                            Set frmClientesNuevo = Nothing
                            
                            Exit Sub
                        Else
                            Unload frmClientesNuevo
                            Set frmClientesNuevo = Nothing
                            
                            Exit Sub
                        End If
                    Else
                        vbq = MsgBox("El nombre ya está registrado como proveedor, ¿Desea convertirlo también en cliente?", vbQuestion + vbYesNo, "Información")
                        If vbq = vbYes Then
                            With Rs
                                If .State = 1 Then .Close
                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                .Open "Select min(id) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(proveedor,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                .Requery
                                In1 = .Fields(0).Value
                                .Close
                            End With
                            sql = "update HZ_PARTY set cliente = 'Si', [Lista de Precios] = " & Combo1 & ",Mayorista ='" & Combo2 & "',last_updated_by ='" & StUsuario & "', last_update_date = '" & Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS") & "' where id = " & In1
                            With Cn
                                .Execute sql
                            End With
                            
                            If InTipoAltaClienteProveedor = 1 Then
                                With frmVentas
                                    .Enabled = True
                                End With
                                Unload Me
                                Set frmClientesNuevo = Nothing
                                
                                Exit Sub
                            Else
                                Unload frmClientesNuevo
                                Set frmClientesNuevo = Nothing
                                
                                Exit Sub
                            End If
                        Else
                            With Text1(0)
                                .SetFocus
                            End With
                            
                            Exit Sub
                        End If
                    End If
                Else
                    MsgBox "El cliente ya esta registrado", vbCritical, "Error"
                    With Text1(0)
                        .SetFocus
                    End With
                    
                    Exit Sub
                End If
                
                With Cn
                    If .State = 1 Then .Close
                End With
            Else
                MsgBox "El nombre es obligatorio", vbCritical, "Error"
                With Text1(0)
                    .SetFocus
                End With
                
                Exit Sub
            End If
        Else
            Unload Me
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesNuevo:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesNuevo:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
