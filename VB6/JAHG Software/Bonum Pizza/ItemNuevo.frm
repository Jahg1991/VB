VERSION 5.00
Begin VB.Form frmItemNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "A�adir Nuevo Articulo"
   ClientHeight    =   3750
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   10410
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
   ScaleHeight     =   3750
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   9855
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
            Left            =   2400
            TabIndex        =   3
            Top             =   1200
            Width           =   2895
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2640
            Width           =   2895
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
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2160
            Width           =   2895
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
            Left            =   2400
            TabIndex        =   4
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   2400
            TabIndex        =   2
            Top             =   720
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   2400
            TabIndex        =   1
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio con IVA"
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
            Left            =   -360
            TabIndex        =   13
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
            Left            =   -360
            TabIndex        =   12
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Iva"
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
            Left            =   -360
            TabIndex        =   11
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio sin IVA"
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
            Left            =   -360
            TabIndex        =   10
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripci�n"
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
            Left            =   -360
            TabIndex        =   9
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C�digo"
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
            Left            =   -360
            TabIndex        =   8
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
Dim Rs As New ADODB.Recordset
Dim NextNumber As Integer

Private Sub Form_Load()
    
    Dim i As Integer
    
    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs
        If .State = 1 Then .Close
        .Open "Select count(*) as folio from items where tipo like '" & StTipoItem & "';", Cn, adOpenStatic, adLockOptimistic
        .Requery
        NextNumber = .Fields(0).Value + 1
        If .State = 1 Then .Close
    End With
    
    Text1(0) = StSerieItem & NextNumber
    
    Text1(1).BackColor = &HC0C0FF
    Text1(3).BackColor = &HC0C0FF
    
    For i = 1 To 2
        Combo1(i).BackColor = &HC0C0FF
    Next i
    
    With Combo1(1)
        .AddItem "0"
        .AddItem "0.16"
        .Text = "0"
    End With
    
    With Combo1(2)
        .AddItem "Chico"
        .AddItem "Familiar"
        .AddItem "Grande"
        .AddItem "Kilogramo"
        .AddItem "Litro"
        .AddItem "Mediano"
        .AddItem "Orden"
        .AddItem "Personal"
        .AddItem "Paquete"
        .AddItem "Pieza"
        .AddItem "Servicio"
        .Text = "Pieza"
    End With
    
    Exit Sub

End Sub

Private Sub Combo1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
        
        Case 1
            
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
            
            Text1(2) = Replace(Format(Val(Text1(3)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
        
        Case 2
            
            With Combo1(2)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
    
    End Select

End Sub

Private Sub Text1_Change(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 1
            
            With Text1(1)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 3
            
            With Text1(3)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
            
            Text1(2) = Replace(Format(Val(Text1(3)) / (1 + Val(Combo1(1))), "0.00"), ",", ".")
    
    End Select

End Sub

Private Sub Guardar_Click()
    
    On Error Resume Next
    
    Dim In1 As Integer
    Dim i As Integer
    
    If Text1(0) <> "" And Text1(1) <> "" And Text1(2) <> "" Then
        
        With Rs
            If .State = 1 Then .Close
            .Open "Select count(*) as existe from items where codigo like '" & Text1(0) & "';", Cn, adOpenStatic, adLockOptimistic
            .Requery
            In1 = .Fields(0).Value
            .Close
        End With
        
        If In1 = 0 Then
        
            With Rs
                If .State = 1 Then .Close
                .Open "Select count(*) as existe from items where descripcion like '" & Text1(1) & "' and UDM like '" & Combo1(2) & "';", Cn, adOpenStatic, adLockOptimistic
                .Requery
                In1 = .Fields(0).Value
                .Close
            End With
            
            If In1 = 0 Then
            
                With Rs
                    If .State = 1 Then .Close
                    .Open "Select * from items;", Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    .AddNew
                        .Fields(1).Value = Text1(0)
                        .Fields(2).Value = Text1(1)
                        .Fields(3).Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                        .Fields(4).Value = Combo1(1)
                        .Fields(5).Value = Combo1(2)
                        .Fields(6).Value = StTipoItem
                    .Update
                    .Requery
                    .Close
                    
                    If .State = 1 Then .Close
                    .Open "Select count(*) as folio from items where tipo like '" & StTipoItem & "';", Cn, adOpenStatic, adLockOptimistic
                    .Requery
                    NextNumber = .Fields(0).Value + 1
                    .Close
                
                End With
                
                Unload frmItemNuevo
                Set frmItemNuevo = Nothing
                frmItemNuevo.Show
                Exit Sub
                
            Else
                MsgBox "La combinaci�n descripci�n y UDM ya existe", vbCritical, "Error"
            End If
        
        Else
            MsgBox "El c�digo esta siendo utilizado por otro art�culo", vbCritical, "Error"
        End If
    
    Else
        MsgBox "Llenar todos los campos", vbCritical, "Error"
    End If

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    frmMenuInicial.Enabled = True
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Sub
