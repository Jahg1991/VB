VERSION 5.00
Begin VB.Form frmProveedoresNuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "A�adir Proveedores"
   ClientHeight    =   5640
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
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   5415
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10140
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   5175
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9855
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   8
            Left            =   2400
            TabIndex        =   19
            Top             =   4560
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   9
            Left            =   8160
            TabIndex        =   20
            Top             =   4560
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   4080
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   7
            Left            =   2400
            TabIndex        =   17
            Top             =   3600
            Width           =   7215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   6
            Left            =   2400
            TabIndex        =   16
            Top             =   3120
            Width           =   7215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   5
            Left            =   2400
            TabIndex        =   15
            Top             =   2640
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   4
            Left            =   2400
            TabIndex        =   14
            Top             =   2160
            Width           =   7215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   3
            Left            =   2400
            TabIndex        =   13
            Top             =   1680
            Width           =   7215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   2
            Left            =   2400
            TabIndex        =   12
            Top             =   1200
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   1
            Left            =   2400
            TabIndex        =   11
            Top             =   720
            Width           =   7215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   0
            Left            =   2400
            TabIndex        =   10
            Top             =   240
            Width           =   7215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de cr�dito $"
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
            TabIndex        =   23
            Top             =   4560
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Limite de cr�dito (d�as)"
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
            Left            =   5280
            TabIndex        =   22
            Top             =   4560
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
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
            Left            =   480
            TabIndex        =   21
            Top             =   4080
            Width           =   1815
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
            Left            =   480
            TabIndex        =   9
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Correo Electronico"
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
            TabIndex        =   8
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tel�fono"
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
            Left            =   840
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
            Left            =   600
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
            Left            =   840
            TabIndex        =   5
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N�mero"
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
            Left            =   840
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
            Left            =   840
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
            Left            =   840
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
Attribute VB_Name = "frmProveedoresNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs      As New adodb.Recordset
    Dim Rs1     As New adodb.Recordset
    
    '//OTROS
    Dim i       As Long
    Dim In1     As Long
    Dim vbq     As Long
    Dim sql     As String
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            
            If .State = 0 Then
                If .State = 0 Then .Open (StConnection)
            End If
        End With
            
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select * from HZ_PARTY_CATEGORIES", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .MoveFirst
            
            While Not .EOF
                frmProveedoresNuevo.Combo1.AddItem .Fields(0).Value
                .MoveNext
            Wend
            
            .Close
        End With
        
        Combo1.Text = "Otros"
        
        For i = 0 To 9
            Text1(i).BackColor = COLOR_NO_ENCONTRADO
        Next i
        
        Text1(8).Text = "0"
        Text1(9).Text = "0"
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmProveedoresNuevo:Form_Load" & vbTab & err.Number & vbTab & err.Description
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
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmProveedoresNuevo:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        vbq = MsgBox("�Desea guardar la informaci�n?", vbQuestion + vbYesNo, "Informaci�n")
                    
        If vbq = vbYes Then
            If Text1(0) <> "" And Text1(5) <> "" Then
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
                        .Open "Select count(*) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(cliente,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
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
                                .Fields(1).Value = Text1(0)
                                .Fields(2).Value = Text1(1)
                                .Fields(3).Value = Text1(2)
                                .Fields(4).Value = Text1(3)
                                .Fields(5).Value = Text1(4)
                                .Fields(6).Value = Text1(5)
                                .Fields(7).Value = Text1(6)
                                .Fields(8).Value = Text1(7)
                                .Fields(9).Value = "Proveedor"
                                .Fields(11).Value = Text1(8)
                                .Fields(12).Value = Text1(9)
                                .Fields(14).Value = Combo1
                                .Fields(16).Value = "Si"
                            .Update
                            .Requery
                            .Close
                        End With
                        
                        If InTipoAltaClienteProveedor = 1 Then
                            frmCompras.Enabled = True
                            
                            Unload Me
                            
                            Set frmProveedoresNuevo = Nothing
                            
                            Exit Sub
                        Else
                            Unload frmProveedoresNuevo
                            
                            Set frmProveedoresNuevo = Nothing
                            
                            frmProveedoresNuevo.Show
                            
                            Exit Sub
                        End If
                    Else
                        vbq = MsgBox("El nombre ya est� registrado como cliente, �Desea convertirlo tambi�n en proveedor?", vbQuestion + vbYesNo, "Informaci�n")
                        
                        If vbq = vbYes Then
                            With Rs
                                If .State = 1 Then .Close
                                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                                .Open "Select min(id) as existe from HZ_PARTY where nombre like '" & Text1(0) & "' and isnull(cliente,'No') = 'Si';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                                .Requery
                                
                                In1 = .Fields(0).Value
                                
                                .Close
                            End With
                            
                            sql = "update HZ_PARTY set proveedor = 'Si' where id = " & In1
                            
                            Cn.Execute sql
                            
                            If InTipoAltaClienteProveedor = 1 Then
                                frmCompras.Enabled = True
                                
                                Unload Me
                                
                                Set frmProveedoresNuevo = Nothing
                                
                                Exit Sub
                            Else
                                Unload frmProveedoresNuevo
                                
                                Set frmProveedoresNuevo = Nothing
                                
                                frmProveedoresNuevo.Show
                                
                                Exit Sub
                            End If
                        Else
                            Text1(0).SetFocus
                        End If
                    End If
                Else
                    MsgBox "El nombre ya existe", vbCritical, "Error"
                    
                    Text1(0).SetFocus
                End If
            Else
                MsgBox "El nombre y el tel�fono son obligatorios", vbCritical, "Error"
                
                Text1(0).SetFocus
            End If
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmProveedoresNuevo:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmProveedoresNuevo:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmProveedoresNuevo:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
