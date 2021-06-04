VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovimientosInventario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Movimientos de inventario"
   ClientHeight    =   7665
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   13965
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
   ScaleHeight     =   7665
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H002B3A4A&
      BorderStyle     =   0  'None
      Caption         =   "HistorialVentasCompras.UDM"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13455
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5775
            Left            =   240
            TabIndex        =   2
            Top             =   960
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   10186
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            HeadLines       =   2
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
            Caption         =   "Movimientos de Inventario"
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
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   0
            Left            =   1200
            TabIndex        =   3
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   118882305
            CurrentDate     =   43915
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Index           =   1
            Left            =   4320
            TabIndex        =   4
            Top             =   120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Format          =   118816769
            CurrentDate     =   43915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Left            =   3480
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmMovimientosInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset

Private Sub Form_Load()
    
    Dim i As Integer
    
    On Error Resume Next
    
    DTPicker1(0).Value = Date
    DTPicker1(1).Value = Date
    
    With Cn
        .CursorLocation = adUseClient
        .Open StConnection
    End With
    
    With Rs
        
        If .State = 1 Then .Close
        
        If StTipoItem = "Barra" Then
            .Open "Select * from TransaccionesDeInventariov where tipo = 'Barra' order by 1,2;", Cn, adOpenStatic, adLockOptimistic
        End If
        
        If StTipoItem = "Cocina" Then
            .Open "Select * from TransaccionesDeInventariov where tipo = 'Cocina' order by 1,2;", Cn, adOpenStatic, adLockOptimistic
        End If
        
        If StTipoItem = "Otros" Then
            .Open "Select * from TransaccionesDeInventariov where (tipo = 'Otros' or tipo = 'Ingredientes Generales') order by 1,2;", Cn, adOpenStatic, adLockOptimistic
        End If
        
        .Requery
        
    End With
    
    If Rs.RecordCount > 0 Then
    
        Rs.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
        Rs.Requery
    
        Set DataGrid1.DataSource = Rs
        
        With DataGrid1
            
            .Columns(0).Width = 1500
            .Columns(1).Width = 1000
            .Columns(2).Width = 4000
            .Columns(3).Width = 3000
            .Columns(4).Width = 1300
            .Columns(5).Width = 1200
            
            .Columns(0).Locked = True
            .Columns(1).Locked = True
            .Columns(2).Locked = True
            .Columns(3).Locked = True
            .Columns(4).Locked = True
            .Columns(5).Locked = True
            
            .Columns(6).Visible = False
        
        End With
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Información"
        frmMenuInicial.Enabled = True
        Unload Me
        
    End If
    
End Sub

Private Sub DTPicker1_Change(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            
            With Rs
        
                .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                .Requery
                
            End With
            
            With DataGrid1
        
                .Columns(0).Width = 1500
                .Columns(1).Width = 1000
                .Columns(2).Width = 4000
                .Columns(3).Width = 3000
                .Columns(4).Width = 1300
                .Columns(5).Width = 1200
                
                .Columns(0).Locked = True
                .Columns(1).Locked = True
                .Columns(2).Locked = True
                .Columns(3).Locked = True
                .Columns(4).Locked = True
                .Columns(5).Locked = True
                
                .Columns(6).Visible = False
            
            End With
            
        Case 1
            
            With Rs
        
                .Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                .Requery
                
            End With
            
            With DataGrid1
        
                .Columns(0).Width = 1500
                .Columns(1).Width = 1000
                .Columns(2).Width = 4000
                .Columns(3).Width = 3000
                .Columns(4).Width = 1300
                .Columns(5).Width = 1200
                
                .Columns(0).Locked = True
                .Columns(1).Locked = True
                .Columns(2).Locked = True
                .Columns(3).Locked = True
                .Columns(4).Locked = True
                .Columns(5).Locked = True
                
                .Columns(6).Visible = False
            
            End With
            
    End Select

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

