VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUsuariosExistente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Usuario"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   690
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   3
      Left            =   12000
      Picture         =   "UsuariosExistente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   2
      Left            =   8400
      Picture         =   "UsuariosExistente.frx":07C8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   1
      Left            =   4200
      Picture         =   "UsuariosExistente.frx":1064
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   0
      Left            =   480
      Picture         =   "UsuariosExistente.frx":1892
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   13455
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5415
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   9551
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
            Caption         =   "Listado de Usuarios"
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
Attribute VB_Name = "frmUsuariosExistente"
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
            .Open "Select * from Usuarios order by 2;", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    
    If Rs.RecordCount > 0 Then
    
        With DataGrid1
            
            Set .DataSource = Rs
            
            For i = 0 To 8
                .Columns(i).Width = 1500
            Next i
            
            .Columns(0).Visible = False
            .Columns(1).Locked = True
        
        End With
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Informaci�n"
        frmMenuInicial.Enabled = True
        Unload Me
        
    End If
    
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

    Dim i As Integer
    
    On Error Resume Next
    
    With Rs
        .Update
        .Requery
    End With
    
    With DataGrid1
        
        For i = 0 To 8
            .Columns(i).Width = 1500
        Next i
        
        .Columns(0).Visible = False
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
