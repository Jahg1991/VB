VERSION 5.00
Begin VB.Form frmItemNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Añadir Nuevo Articulo"
   ClientHeight    =   3960
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
   ScaleHeight     =   3960
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3495
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9855
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   0
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2640
            Width           =   2895
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   2
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2160
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
            TabIndex        =   10
            Top             =   1680
            Width           =   2895
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
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Index           =   1
            Left            =   2400
            TabIndex        =   8
            Top             =   720
            Width           =   7335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   2400
            TabIndex        =   7
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Categoría"
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
            TabIndex        =   12
            Top             =   2640
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
            TabIndex        =   6
            Top             =   2160
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
            TabIndex        =   5
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
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
            TabIndex        =   4
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
            Left            =   -360
            TabIndex        =   3
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
            Left            =   -360
            TabIndex        =   2
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

Private Sub Form_Load()
    
    Dim i As Integer
    
    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    For i = 0 To 2
        Text1(i).BackColor = &HC0C0FF
        Combo1(i).BackColor = &HC0C0FF
    Next i
    
    With Combo1(0)
        .AddItem "Res"
        .AddItem "Pollo"
        .AddItem "Cerdo"
        .AddItem "Otro"
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
    End With
    
    Exit Sub

End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            With Combo1(0)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 1
            
            With Combo1(1)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
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
        
        Case 0
            
            With Text1(0)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 1
            
            With Text1(1)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
        
        Case 2
            
            With Text1(2)
                If .Text = "" Then
                    .BackColor = &HC0C0FF
                Else
                    .BackColor = &HE0E0E0
                End If
            End With
    
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
                .Open "Select * from items;", Cn, adOpenStatic, adLockOptimistic
                .Requery
                .AddNew
                    .Fields(1).Value = Text1(0)
                    .Fields(2).Value = Text1(1)
                    .Fields(3).Value = Replace(Format(Val(Text1(2)), "0.00"), ",", ".")
                    .Fields(4).Value = Combo1(1)
                    .Fields(5).Value = Combo1(2)
                    .Fields(6).Value = Combo1(0)
                .Update
                .Requery
                .Close
            
            End With
            
            Unload frmItemNuevo
            Set frmItemNuevo = Nothing
            frmItemNuevo.Show
            Exit Sub
        
        Else
            MsgBox "El código esta siendo utilizado por otro artículo", vbCritical, "Error"
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
