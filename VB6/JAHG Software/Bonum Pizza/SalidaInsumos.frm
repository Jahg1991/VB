VERSION 5.00
Begin VB.Form frmSalidaInsumos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Insumos"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1560
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1335
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   11055
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Top             =   120
            Width           =   9375
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   3
            Top             =   600
            Width           =   5175
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
            Index           =   12
            Left            =   -600
            TabIndex        =   5
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
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
            Left            =   -600
            TabIndex        =   4
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
Attribute VB_Name = "frmSalidaInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset

Dim InItemId As Integer

Dim v1 As Integer
Dim v2 As String
Dim v3 As String
Dim v4 As String
Dim v5 As String
Dim v6 As String
Dim v7 As String
Dim v8 As String
Dim v9 As String

Private Sub Form_Load()

    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    List1.Clear
    List2.Clear
    
    With Rs
        
        If .State = 1 Then .Close
        
        .Open "Select * from items where udm <> 'Servicio' order by 3;", Cn, adOpenStatic, adLockOptimistic
        .Filter = ""
        .Requery
        
        If .RecordCount <> 0 Then
            
            Combo1.Clear
            
            While Not .EOF
                Combo1.AddItem .Fields(2) & " (" & .Fields(5) & ")"
                .MoveNext
            Wend
        
        End If
    
    End With
    
    Combo1.BackColor = &HC0C0FF
    Text1.BackColor = &HC0C0FF
    
End Sub

Private Sub Combo1_Click()

    On Error Resume Next
    
    With Combo1
                
        If .Text = "" Then
                    
            .BackColor = &HC0C0FF
                    
            With Rs
                .Filter = ""
                .Requery
            End With
                
        Else
                
            .BackColor = &HE0E0E0
                    
            InItemId = Get_ItemId(.Text)
                    
            With Rs
                .Filter = "Id = " & InItemId
                .Requery
            End With
                
        End If
        
    End With
    
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    With Combo1
                
        ' si pesionamos las teclas de las flechas sale de la rutina
        If KeyCode = vbKeyUp Then Exit Sub
        If KeyCode = vbKeyDown Then Exit Sub
        If KeyCode = vbKeyLeft Then Exit Sub
        If KeyCode = vbKeyRight Then Exit Sub
                
        ' verifica qu no se presionó la tecla backspace
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
            .BackColor = &HC0C0FF
                    
            With Rs
                .Filter = ""
                .Requery
            End With
                
        Else
                    
            ' Backcolor normal cuando hay coincidencia
            .BackColor = &HE0E0E0
                    
            InItemId = Get_ItemId(.Text)
                    
            With Rs
                .Filter = "Id = " & InItemId
                .Requery
            End With
                
        End If
            
    End With
            
End Sub

Private Sub Text1_Change()
    
    On Error Resume Next
    
    With Text1
        If .Text = "" Then
            .BackColor = &HC0C0FF
        Else
            .BackColor = &HE0E0E0
        End If
    End With
            
End Sub

Private Sub Guardar_Click()

    On Error Resume Next
    
        With Rs1
            
            If .State = 1 Then .Close
            
            .Open "Select * from TransaccionesDeInventario;", Cn, adOpenStatic, adLockOptimistic
            .Filter = ""
            .Requery
            
            If Combo1 <> "" And Text1 <> "" Then
            
                v1 = InItemId
                v2 = Get_ItemCod(v1)
                v3 = Get_ItemDesc(v1)
                v4 = Date
                v6 = Replace(Format(Val(Text1) * -1, "0.00"), ",", ".")
                v5 = "Salida Insumos"
                v7 = Get_ItemUDM(v1)
                v8 = ""
                v9 = "No"
                        
                .AddNew
                    .Fields(1) = v1
                    .Fields(2) = v2
                    .Fields(3) = v3
                    .Fields(4) = v4
                    .Fields(5) = v5
                    .Fields(6) = v6
                    .Fields(7) = v7
                    .Fields(8) = v8
                    .Fields(9) = v9
                .Update
                .Requery
                
                MsgBox "Salida de Insumos Registrada", vbOKOnly, "Información"
            
            Else
            
                MsgBox "Llene los campos artículo y cantidad", vbCritical, "Advertencia"
                
                Exit Sub
                
            End If
        
        End With
    
        Unload frmSalidaInsumos
        Set frmSalidaInsumos = Nothing
        frmSalidaInsumos.Show

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next

    frmMenuInicial.Enabled = True
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Rs1.State = 1 Then Rs.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Rs1 = Nothing
    Set Cn = Nothing

End Sub
