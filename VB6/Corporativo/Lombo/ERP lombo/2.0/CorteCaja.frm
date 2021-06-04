VERSION 5.00
Begin VB.Form frmCorteCaja 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Corte de Caja"
   ClientHeight    =   3615
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   6930
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
   ScaleHeight     =   3615
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6660
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6375
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   5
            Left            =   2880
            TabIndex        =   7
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   4
            Left            =   2880
            TabIndex        =   6
            Top             =   2040
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   3
            Left            =   2880
            TabIndex        =   5
            Top             =   1560
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   2
            Left            =   2880
            TabIndex        =   4
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   1
            Left            =   2880
            TabIndex        =   3
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   0
            Left            =   2880
            TabIndex        =   2
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Efectivo final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Retiro de Efectivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada de efectivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Efectivo Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   2655
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmCorteCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset
Dim Rs5 As New ADODB.Recordset

Private Sub Form_Load()
    
    On Error Resume Next
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs
        
        If .State = 1 Then .Close
        .Open "Select sum(cantidad) from MovimientosCajaV where fecha < #" & Date & "#;", Cn, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    With Rs1
        
        If .State = 1 Then .Close
        .Open "Select cantidad from MovimientosCajaV where fecha = #" & Date & "# and TipoMovimiento = 'Pago de venta';", Cn, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    With Rs2
        
        If .State = 1 Then .Close
        .Open "Select cantidad from MovimientosCajaV where fecha = #" & Date & "# and TipoMovimiento = 'Pago de compra';", Cn, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    With Rs3
        
        If .State = 1 Then .Close
        .Open "Select cantidad from MovimientosCajaV where fecha = #" & Date & "# and TipoMovimiento = 'Entrada Manual de Efectivo';", Cn, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    With Rs4
        
        If .State = 1 Then .Close
        .Open "Select cantidad from MovimientosCajaV where fecha = #" & Date & "# and TipoMovimiento = 'Retiro Manual de Efectivo';", Cn, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    With Rs5
        
        If .State = 1 Then .Close
        .Open "Select sum(cantidad) from MovimientosCajaV where fecha <= #" & Date & "#;", Cn, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    Text2(0) = Replace(Format(Rs.Fields(0).Value, "0.00"), ",", ".")
    Text2(1) = Replace(Format(Rs1.Fields(0).Value, "0.00"), ",", ".")
    Text2(2) = Replace(Format(Rs2.Fields(0).Value, "0.00"), ",", ".")
    Text2(3) = Replace(Format(Rs3.Fields(0).Value, "0.00"), ",", ".")
    Text2(4) = Replace(Format(Rs4.Fields(0).Value, "0.00"), ",", ".")
    Text2(5) = Replace(Format(Rs5.Fields(0).Value, "0.00"), ",", ".")
    
End Sub

Private Sub Imprimir_Click()
    
    On Error Resume Next
    
    Unload TicketCorte
    
    With TicketCorte
    
        Set .DataSource = Rs
        
        With .Sections("Sección4")
            .Controls("Etiqueta2").Caption = PcNombreEmpresa
            .Controls("Etiqueta10").Caption = Text2(0)
            .Controls("Etiqueta11").Caption = Text2(1)
            .Controls("Etiqueta12").Caption = Text2(2)
            .Controls("Etiqueta13").Caption = Text2(3)
            .Controls("Etiqueta14").Caption = Text2(4)
            .Controls("Etiqueta15").Caption = Text2(5)
        End With
    
        .Hide
        
        .PrintReport True
        
    End With
    
    Unload TicketCorte
    
End Sub

Private Sub Salir_Click()
    
    On Error Resume Next

    frmMenuInicial.Enabled = True
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    If Rs.State = 1 Then Rs.Close
    If Rs1.State = 1 Then Rs1.Close
    If Rs2.State = 1 Then Rs2.Close
    If Rs3.State = 1 Then Rs3.Close
    If Rs4.State = 1 Then Rs4.Close
    If Rs5.State = 1 Then Rs5.Close
    If Cn.State = 1 Then Cn.Close
    
    Set Rs = Nothing
    Set Rs1 = Nothing
    Set Rs2 = Nothing
    Set Rs3 = Nothing
    Set Rs4 = Nothing
    Set Rs5 = Nothing
    Set Cn = Nothing

End Sub
