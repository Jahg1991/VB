VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmStock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Stock"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
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
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   0
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   11655
         End
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
            Caption         =   "Stock"
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
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmStock"
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
        .Open "Select * from stock order by 3;", Cn, adOpenStatic, adLockOptimistic
        .Requery
        
    End With
    
    If Rs.RecordCount > 0 Then
    
        Set DataGrid1.DataSource = Rs
        
        With DataGrid1
            
            .Columns(2).Width = 2000
            .Columns(3).Width = 5000
            .Columns(4).Width = 3000
            .Columns(5).Width = 2000
            
            .Columns(0).Locked = True
            .Columns(1).Locked = True
            .Columns(2).Locked = True
            .Columns(3).Locked = True
            .Columns(4).Locked = True
            .Columns(5).Locked = True
            
            .Columns(0).Visible = False
            .Columns(1).Visible = False
        
        End With
        
    Else
    
        MsgBox "No hay registros existentes", vbOKOnly, "Informaci?n"
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
                    .Filter = "ItemCodigo like '*" & Text1(0) & "*' or ItemDescricion like '*" & Text1(0) & "*'"
                    .Requery
                End If
            End With
            
            With DataGrid1
        
                .Columns(2).Width = 2000
                .Columns(3).Width = 5000
                .Columns(4).Width = 3000
                .Columns(5).Width = 2000
                
                .Columns(0).Locked = True
                .Columns(1).Locked = True
                .Columns(2).Locked = True
                .Columns(3).Locked = True
                .Columns(4).Locked = True
                .Columns(5).Locked = True
                
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            
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
