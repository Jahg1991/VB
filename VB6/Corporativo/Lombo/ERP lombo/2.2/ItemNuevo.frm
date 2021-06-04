VERSION 5.00
Begin VB.Form frmItemNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Añadir Nuevo Articulo"
   ClientHeight    =   5805
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   11220
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5535
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5295
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   10695
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   3
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   4560
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   7920
            TabIndex        =   16
            Top             =   4080
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            Height          =   420
            Index           =   11
            Left            =   7920
            TabIndex        =   12
            Top             =   3120
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            Height          =   420
            Index           =   10
            Left            =   7920
            TabIndex        =   10
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            Height          =   420
            Index           =   9
            Left            =   7920
            TabIndex        =   8
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            Height          =   420
            Index           =   8
            Left            =   7920
            TabIndex        =   6
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Height          =   420
            Index           =   7
            Left            =   2760
            TabIndex        =   11
            Top             =   3120
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Height          =   420
            Index           =   6
            Left            =   2760
            TabIndex        =   9
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Height          =   420
            Index           =   5
            Left            =   2760
            TabIndex        =   7
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Height          =   420
            Index           =   4
            Left            =   2760
            TabIndex        =   5
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   0
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   4080
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            Height          =   420
            Index           =   3
            Left            =   2760
            TabIndex        =   3
            Top             =   1200
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   7920
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3600
            Width           =   2535
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
            Height          =   420
            Index           =   1
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   3600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            Enabled         =   0   'False
            Height          =   420
            Index           =   2
            Left            =   7920
            TabIndex        =   4
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   2760
            TabIndex        =   2
            Top             =   720
            Width           =   7695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   0
            Left            =   2760
            TabIndex        =   1
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Artículo"
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
            Index           =   16
            Left            =   0
            TabIndex        =   35
            Top             =   4560
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Control de lote"
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
            Index           =   15
            Left            =   5040
            TabIndex        =   34
            Top             =   4080
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 2 sin IVA"
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
            Index           =   14
            Left            =   5160
            TabIndex        =   33
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 3 sin IVA"
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
            Index           =   13
            Left            =   5160
            TabIndex        =   32
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 4 sin IVA"
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
            Left            =   5160
            TabIndex        =   31
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 5 (C) sin IVA"
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
            TabIndex        =   30
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 2 con IVA"
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
            Left            =   0
            TabIndex        =   29
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 3 con IVA"
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
            Left            =   0
            TabIndex        =   28
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 4 con IVA"
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
            Left            =   0
            TabIndex        =   27
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 5 (C) con IVA"
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
            Left            =   120
            TabIndex        =   26
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación"
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
            Left            =   0
            TabIndex        =   25
            Top             =   4080
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 1 con IVA"
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
            Left            =   0
            TabIndex        =   24
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad de medida"
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
            Left            =   5160
            TabIndex        =   23
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa IVA"
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
            Left            =   0
            TabIndex        =   22
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio 1 sin IVA"
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
            Left            =   5160
            TabIndex        =   21
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Left            =   0
            TabIndex        =   20
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
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
            Left            =   0
            TabIndex        =   19
            Top             =   240
            Width           =   2655
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
Attribute VB_Name = "frmItemNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs  As New adodb.Recordset
    Dim Rs1 As New adodb.Recordset
    
    '//OTROS
    Dim i   As Long
    Dim In1 As Long

    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        Text1(0).BackColor = COLOR_NO_ENCONTRADO
        Text1(1).BackColor = COLOR_NO_ENCONTRADO
        Text1(3).BackColor = COLOR_NO_ENCONTRADO
        Text1(4).BackColor = COLOR_NO_ENCONTRADO
        Text1(5).BackColor = COLOR_NO_ENCONTRADO
        Text1(6).BackColor = COLOR_NO_ENCONTRADO
        Text1(7).BackColor = COLOR_NO_ENCONTRADO
        
        For i = 0 To 3
            Combo1(i).BackColor = COLOR_NORMAL
        Next i
        
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select Categoria From MTL_ITEM_CATEGORIES order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            .MoveFirst
            
            While Not .EOF
                frmItemNuevo.Combo1(0).AddItem .Fields(0).Value
                
                .MoveNext
            Wend
            
            .MoveFirst
            
            frmItemNuevo.Combo1(0).Text = .Fields(0).Value
            
            .Close
        End With
        
        With Combo1(1)
            .AddItem "0"
            .AddItem "0.16"
            
            .Text = "0"
        End With
        
        With Combo1(2)
            .AddItem "Kilogramo"
            .AddItem "Litro"
            .AddItem "Pieza"
            .AddItem "Servicio"
            
            .Text = "Kilogramo"
        End With
        
        With Combo1(3)
            .AddItem "Inventario"
            .AddItem "Gasto"
            
            .Text = "Inventario"
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                With Combo1(2)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
            
            Case 1
                With Combo1(1)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
                Text1(2) = Replace(Format(Val(Text1(3)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
                Text1(8) = Replace(Format(Val(Text1(4)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
                Text1(9) = Replace(Format(Val(Text1(5)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
                Text1(10) = Replace(Format(Val(Text1(6)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
                Text1(11) = Replace(Format(Val(Text1(7)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
            
            Case 2
                With Combo1(2)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
            Case 3
                With Combo1(3)
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
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
            
            Case 3
                With Text1(3)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
                Text1(2) = Replace(Format(Val(Text1(3)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
            
            Case 4
                With Text1(4)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
                Text1(8) = Replace(Format(Val(Text1(4)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
                
            Case 5
                With Text1(5)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
                Text1(9) = Replace(Format(Val(Text1(5)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
            
            Case 6
                With Text1(6)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
                Text1(10) = Replace(Format(Val(Text1(6)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
                
            Case 7
                With Text1(7)
                    If .Text = "" Then
                        .BackColor = COLOR_NO_ENCONTRADO
                    Else
                        .BackColor = COLOR_NORMAL
                    End If
                End With
                
                Text1(11) = Replace(Format(Val(Text1(7)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
                    
        If vbq = vbYes Then
            If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" And Combo1(3) <> "" Then
                With Rs
                    If .State = 1 Then .Close
                    .CursorLocation = adodb.CursorLocationEnum.adUseClient
                    .Open "Select count(*) as existe from MTL_SYSTEM_ITEMS where codigo like '" & Text1(0) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                    .Requery
                    
                    In1 = .Fields(0).Value
                    
                    .Close
                End With
                
                If In1 = 0 Then
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select count(*) as existe from MTL_SYSTEM_ITEMS where descripcion like '" & Text1(1) & "' and UDM like '" & Combo1(2) & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        
                        In1 = .Fields(0).Value
                        
                        .Close
                    End With
                    
                    If In1 = 0 Then
                        With Rs
                            If .State = 1 Then .Close
                            .CursorLocation = adodb.CursorLocationEnum.adUseClient
                            .Open "Select * from MTL_SYSTEM_ITEMS;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                            .Requery
                            .AddNew
                                .Fields(1).Value = Text1(0)
                                .Fields(2).Value = Text1(1)
                                .Fields(3).Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                                .Fields(4).Value = Replace(Format(Val(Text1(8)), "0.00"), ",", ".")
                                .Fields(5).Value = Replace(Format(Val(Text1(9)), "0.00"), ",", ".")
                                .Fields(6).Value = Replace(Format(Val(Text1(10)), "0.00"), ",", ".")
                                .Fields(7).Value = Replace(Format(Val(Text1(11)), "0.00"), ",", ".")
                                .Fields(8).Value = Combo1(1)
                                .Fields(9).Value = Combo1(2)
                                .Fields(10).Value = Combo1(0)
                                .Fields(11).Value = Check1
                                .Fields(12).Value = Combo1(3)
                            .Update
                            .Requery
                            .Close
                        End With
                        
                        Unload frmItemNuevo
                        
                        Set frmItemNuevo = Nothing
                        
                        frmItemNuevo.Show
                        
                        Exit Sub
                    Else
                        MsgBox "La combinación Descripción - UDM ya existe", vbCritical, "Error"
                    End If
                Else
                    MsgBox "El código esta siendo utilizado por otro artículo", vbCritical, "Error"
                End If
            Else
                MsgBox "Llenar todos los campos", vbCritical, "Error"
            End If
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmItemNuevo:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
