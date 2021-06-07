VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
         TabIndex        =   1
         Top             =   120
         Width           =   16935
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4440
            TabIndex        =   11
            Top             =   2640
            Width           =   255
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
            Index           =   4
            Left            =   4440
            TabIndex        =   10
            Top             =   2160
            Width           =   4095
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
            Left            =   4440
            TabIndex        =   9
            Top             =   1680
            Width           =   4095
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
            Left            =   4440
            TabIndex        =   8
            Top             =   1200
            Width           =   12375
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
            Left            =   4440
            TabIndex        =   7
            Top             =   720
            Width           =   3975
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
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   12375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "INVENTARIOS NEGATIVOS"
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
            Left            =   0
            TabIndex        =   13
            Top             =   2640
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NO. DE PUNTOS POR CADA PESO"
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
            Left            =   0
            TabIndex        =   12
            Top             =   2160
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RFC"
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
            Left            =   2160
            TabIndex        =   5
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DIRECCION"
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
            Left            =   2160
            TabIndex        =   4
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONO"
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
            Left            =   2160
            TabIndex        =   3
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NOMBRE EMPRESA"
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
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   3855
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
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmOpciones
'Proposito:     Define preferencias del sistema
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'1.1        14/05/2021     Alfredo Hernandez    Agrego manejo de errores para error
'                                               3219. Cuando hay modificaciones y se
'                                               sale sin guardar.
'
'1.2        14/05/2021     Alfredo Hernandez    Se agrego confirmacion de salida sin
'                                               guardar datos
'
'***********************************************************************************
Option Explicit

'===============================================================================
'DECLARACION DE VARIABLES
'===============================================================================

'//RECORDSET
Dim Rs As New adodb.Recordset
'//OTROS
Dim i As Long
Dim ctl As Printer

Private Sub Form_Load()
    On Error GoTo errHandler
    With Cn
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open (StConnection)
    End With

    With Rs
        If .State = 1 Then .Close
        .CursorLocation = adodb.CursorLocationEnum.adUseClient
        .Open "Select * from FND_SYSTEM_OPTIONS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
        .Requery
        .MoveFirst
    End With

    For i = 0 To 4
        With Text1(i)
            Set .DataSource = Rs
            .BackColor = COLOR_NO_ENCONTRADO
        End With
    Next i

    With Text1(0)
        .DataField = "NombreEmpresa"
    End With

    With Text1(1)
        .DataField = "RFC"
    End With

    With Text1(2)
        .DataField = "Direccion"
    End With

    With Text1(3)
        .DataField = "Telefono"
    End With

    With Text1(4)
        .DataField = "ValorPuntos"
    End With

    With Check1
        Set .DataSource = Rs
        .DataField = "Inventarios"
    End With
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmOpciones:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
    End Select
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmOpciones:Text1_Change" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Guardar_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Rs
            .Fields("last_updated_by").Value = StUsuario
            .Fields("last_update_date").Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
            .Update
        End With
        Unload Me
        Unload frmMenuInicial
        Main
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmOpciones:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub

Private Sub Salir_Click()
    On Error GoTo errHandler
    vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
    If vbq = vbYes Then
        With Rs
            .Fields("last_updated_by").Value = StUsuario
            .Fields("last_update_date").Value = Format(Date, "YYYY-MM-DD") & " " & Format(Time, "HH:MM:SS")
            .Update
        End With
        Unload Me
        Unload frmMenuInicial
        Main
    Else
        Unload Me
    End If
    Exit Sub
errHandler:
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmOpciones:Salir_Click" & vbTab & err.Number & vbTab & err.Description
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
    If err.Number = 3219 Then
        Set frmOpciones = Nothing
        Unload Me
        Exit Sub
    End If
    FileNum = FreeFile
    Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
    Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmOpciones:Form_Unload" & vbTab & err.Number & vbTab & err.Description
    Close FileNum
    err.Clear
    MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
End Sub
