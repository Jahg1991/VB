VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClientesProveedoresExistente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Catalogo Clientes / Proveedores"
   ClientHeight    =   7260
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13905
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
   ScaleHeight     =   7260
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   13335
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   11655
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4695
            Left            =   240
            TabIndex        =   2
            Top             =   960
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   8281
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            HeadLines       =   3
            RowHeight       =   28
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   0
            Left            =   240
            Picture         =   "ClientesProveedoresExistente.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   1
            Left            =   3960
            Picture         =   "ClientesProveedoresExistente.frx":07EE
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   2
            Left            =   7920
            Picture         =   "ClientesProveedoresExistente.frx":101C
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5880
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   3
            Left            =   11640
            Picture         =   "ClientesProveedoresExistente.frx":18B8
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   5880
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar"
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
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   975
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
Attribute VB_Name = "frmClientesProveedoresExistente"
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
    
    With Rs
            If .State = 1 Then .Close
            .Open "Select * from ClientesProveedores where tipo = '" & StTipoClienteProveedor & "' order by 2;", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    If Rs.RecordCount > 0 Then
    
        With DataGrid1
            
            If StTipoClienteProveedor = "Cliente" Then
                .Caption = "Listado de Clientes"
            Else
                .Caption = "Listado de Proveedores"
            End If
            
            Set .DataSource = Rs
            
            For i = 0 To 9
                .Columns(i).Width = 1500
            Next i
            
            .Columns(0).Visible = False
            .Columns(9).Visible = False
            .Columns(1).Locked = True
        
        End With
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Información"
        frmMenuInicial.Enabled = True
        Unload Me
    
    End If

End Sub

Private Sub Text1_Change(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            Dim i As Integer
            
            With Rs
                
                If Text1(0) = "" Then
                    .Filter = ""
                    .Requery
                Else
                    .Filter = "nombre like '*" & Text1(0) & "*' or telefono like '*" & Text1(0) & "*'"
                    .Requery
                End If
            
            End With
            
            With DataGrid1
                
                If StTipoClienteProveedor = "Cliente" Then
                    .Caption = "Listado de Clientes"
                Else
                    .Caption = "Listado de Proveedores"
                End If
                
                Set .DataSource = Rs
                
                For i = 0 To 9
                    .Columns(i).Width = 1500
                Next i
                
                .Columns(0).Visible = False
                .Columns(9).Visible = False
                .Columns(1).Locked = True
            
            End With
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
        
        Case 0
            
            Rs.MoveFirst
        
        Case 1
            
            Rs.MovePrevious
        
        Case 2
            
            Rs.MoveNext
        
        Case 3
            
            Rs.MoveLast
    
    End Select

End Sub

Private Sub Guardar_Click()
    
    On Error Resume Next
    
    Dim i As Integer
    
    With Rs
        .Update
        .Requery
    End With
    
    With DataGrid1
        Set .DataSource = Rs
        
        For i = 0 To 9
            .Columns(i).Width = 1500
        Next i
        
        .Columns(0).Visible = False
        .Columns(9).Visible = False
        .Columns(1).Locked = True
    
    End With

End Sub

Private Sub Salir_Click()
    
    On Error Resume Next
    
    Rs.Requery
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
