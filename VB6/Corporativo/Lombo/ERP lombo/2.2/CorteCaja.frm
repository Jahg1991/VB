VERSION 5.00
Begin VB.Form frmCorteCaja 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Corte de Caja"
   ClientHeight    =   4575
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
   ScaleHeight     =   4575
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   4335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6660
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "292_Nutec_29-SEP-20.csv"
         ForeColor       =   &H00FFFFFF&
         Height          =   4095
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
            Index           =   7
            Left            =   2880
            TabIndex        =   17
            Top             =   3480
            Width           =   3255
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   420
            Index           =   6
            Left            =   2880
            TabIndex        =   15
            Top             =   3000
            Width           =   3255
         End
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
            Caption         =   "iZettle"
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
            Index           =   7
            Left            =   120
            TabIndex        =   16
            Top             =   3480
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Banamex"
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
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   3000
            Width           =   2655
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
    Option Explicit
    
    '//RECORDSET
    Dim Rs  As New adodb.Recordset
    Dim Rs1 As New adodb.Recordset
    Dim Rs2 As New adodb.Recordset
    Dim Rs3 As New adodb.Recordset
    Dim Rs4 As New adodb.Recordset
    Dim Rs5 As New adodb.Recordset
    Dim Rs6 As New adodb.Recordset
    Dim Rs7 As New adodb.Recordset
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select sum(cantidad) from RA_CASH_TRANSACTIONS_V where fecha < CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and caja = '" & frmMenuInicial.Combo1.Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs1
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select cantidad from RA_CASH_TRANSACTIONS_V where fecha = CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and TipoMovimiento = 'Pago de venta' and caja = '" & frmMenuInicial.Combo1.Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs2
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select cantidad from RA_CASH_TRANSACTIONS_V where fecha = CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and TipoMovimiento = 'Pago de compra' and caja = '" & frmMenuInicial.Combo1.Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs3
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select cantidad from RA_CASH_TRANSACTIONS_V where fecha = CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and TipoMovimiento = 'Entrada Manual de Efectivo' and caja = '" & frmMenuInicial.Combo1.Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs4
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select cantidad from RA_CASH_TRANSACTIONS_V where fecha = CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and TipoMovimiento = 'Retiro Manual de Efectivo' and caja = '" & frmMenuInicial.Combo1.Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs5
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select sum(cantidad) from RA_CASH_TRANSACTIONS_V where fecha <= CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and caja = '" & frmMenuInicial.Combo1.Text & "';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs6
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select sum(cantidad) from RA_BANK_TRANSACTIONS where fecha = CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and caja = '" & frmMenuInicial.Combo1.Text & "' and cancelado = 'No' and TipoTarjeta = 'Banamex';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        With Rs7
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "Select sum(cantidad) from RA_BANK_TRANSACTIONS where fecha = CONVERT(DATETIME, CONVERT(DATE, GETDATE())) and caja = '" & frmMenuInicial.Combo1.Text & "' and cancelado = 'No' and TipoTarjeta = 'iZettle';", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
        End With
        
        If Rs.RecordCount <> 0 Then Text2(0) = Replace(Format(Rs.Fields(0).Value, "0.00"), ",", ".")
        If Rs1.RecordCount <> 0 Then Text2(1) = Replace(Format(Rs1.Fields(0).Value, "0.00"), ",", ".")
        If Rs2.RecordCount <> 0 Then Text2(2) = Replace(Format(Rs2.Fields(0).Value, "0.00"), ",", ".")
        If Rs3.RecordCount <> 0 Then Text2(3) = Replace(Format(Rs3.Fields(0).Value, "0.00"), ",", ".")
        If Rs4.RecordCount <> 0 Then Text2(4) = Replace(Format(Rs4.Fields(0).Value, "0.00"), ",", ".")
        If Rs5.RecordCount <> 0 Then Text2(5) = Replace(Format(Rs5.Fields(0).Value, "0.00"), ",", ".")
        If Rs6.RecordCount <> 0 Then Text2(6) = Replace(Format(Rs6.Fields(0).Value, "0.00"), ",", ".")
        If Rs7.RecordCount <> 0 Then Text2(7) = Replace(Format(Rs7.Fields(0).Value, "0.00"), ",", ".")
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCorteCaja:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Imprimir_Click()
        On Error GoTo errHandler
        
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
                .Controls("Label3").Caption = Text2(6)
                .Controls("Label5").Caption = Text2(7)
                .Controls("Etiqueta17").Caption = frmMenuInicial.Combo1.Text
            End With
        
            .Show 1
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCorteCaja:Imprimir_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCorteCaja:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        Unload TicketCorte
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Rs2.State = 1 Then Rs2.Close
        If Rs3.State = 1 Then Rs3.Close
        If Rs4.State = 1 Then Rs4.Close
        If Rs5.State = 1 Then Rs5.Close
        If Rs6.State = 1 Then Rs6.Close
        If Rs7.State = 1 Then Rs7.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Rs2 = Nothing
        Set Rs3 = Nothing
        Set Rs4 = Nothing
        Set Rs5 = Nothing
        Set Rs6 = Nothing
        Set Rs7 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCorteCaja:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
