VERSION 5.00
Begin VB.Form frmClientesConfirmacionDatos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Confirmación de Información del cliente"
   ClientHeight    =   5610
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   10335
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
   ScaleHeight     =   5610
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10065
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9855
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   12
            Left            =   6960
            TabIndex        =   20
            Top             =   3600
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   11
            Left            =   2280
            TabIndex        =   19
            Top             =   3600
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   10
            Left            =   6960
            TabIndex        =   18
            Top             =   3120
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   9
            Left            =   2280
            TabIndex        =   17
            Top             =   3120
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   8
            Left            =   6960
            TabIndex        =   16
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   7
            Left            =   2280
            TabIndex        =   22
            Top             =   4560
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   6
            Left            =   2280
            TabIndex        =   21
            Top             =   4080
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   5
            Left            =   2280
            TabIndex        =   15
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   4
            Left            =   2280
            TabIndex        =   14
            Top             =   2160
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   3
            Left            =   2280
            TabIndex        =   13
            Top             =   1680
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   2
            Left            =   2280
            TabIndex        =   12
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   1
            Left            =   2280
            TabIndex        =   11
            Top             =   720
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Width           =   7335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 6"
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
            Index           =   12
            Left            =   5400
            TabIndex        =   27
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 4"
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
            Index           =   11
            Left            =   5400
            TabIndex        =   26
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 5"
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
            Index           =   10
            Left            =   720
            TabIndex        =   25
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 3"
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
            Index           =   9
            Left            =   720
            TabIndex        =   24
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono 2"
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
            Index           =   8
            Left            =   5400
            TabIndex        =   23
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencias"
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
            Index           =   7
            Left            =   360
            TabIndex        =   9
            Top             =   4560
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. tarjeta"
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
            Index           =   6
            Left            =   -120
            TabIndex        =   8
            Top             =   4080
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono"
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
            Left            =   720
            TabIndex        =   7
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Postal"
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
            Left            =   480
            TabIndex        =   6
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Colonia"
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
            Left            =   720
            TabIndex        =   5
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
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
            Left            =   720
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calle"
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
            Left            =   720
            TabIndex        =   3
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
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
            Left            =   720
            TabIndex        =   2
            Top             =   240
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
Attribute VB_Name = "frmClientesConfirmacionDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs10    As New adodb.Recordset
    
    '//OTROS
    Dim i       As Long
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Rs10
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from HZ_PARTY where id =" & IdCliente, Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        For i = 0 To (12)
            Set Text1(i).DataSource = Rs10
        Next i
        
        For i = 1 To (12)
            Text1(i).BackColor = COLOR_NO_ENCONTRADO
        Next i
        
        Text1(0).DataField = "Nombre"
        Text1(1).DataField = "Calle"
        Text1(2).DataField = "Numero"
        Text1(3).DataField = "Colonia"
        Text1(4).DataField = "Codigo Postal"
        Text1(5).DataField = "Telefono"
        Text1(6).DataField = "Correo Electronico"
        Text1(7).DataField = "referencias"
        Text1(8).DataField = "Telefono2"
        Text1(9).DataField = "Telefono3"
        Text1(10).DataField = "Telefono4"
        Text1(11).DataField = "Telefono5"
        Text1(12).DataField = "Telefono6"
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesConfirmacionDatos:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Text1_Change(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
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
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesConfirmacionDatos:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
                    
        If vbq = vbYes Then
            With Rs10
                .Fields(1) = Text1(0)
                .Fields(2) = Text1(1)
                .Fields(3) = Text1(2)
                .Fields(4) = Text1(3)
                .Fields(5) = Text1(4)
                .Fields(6) = Text1(5)
                .Fields(7) = Text1(8)
                .Fields(8) = Text1(9)
                .Fields(9) = Text1(10)
                .Fields(10) = Text1(11)
                .Fields(11) = Text1(12)
                .Fields(12) = Text1(6)
                .Fields(13) = Text1(7)
                .Update
                .Requery
            End With
        End If
        
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesConfirmacionDatos:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesConfirmacionDatos:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs10.State = 1 Then Rs10.Close
        
        Set Rs10 = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmClientesConfirmacionDatos:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
