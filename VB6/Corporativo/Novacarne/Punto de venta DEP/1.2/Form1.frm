VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   14025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22095
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "Arial Narrow"
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
   ScaleHeight     =   14025
   ScaleWidth      =   22095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Frame FramePrincipal 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   13455
      Left            =   3240
      TabIndex        =   9
      Top             =   360
      Width           =   18555
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   10575
         Index           =   1
         Left            =   6840
         TabIndex        =   11
         Top             =   240
         Width           =   11115
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   4
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   6840
            Width           =   2415
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   735
            Left            =   480
            TabIndex        =   30
            Top             =   8640
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   1296
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            BackColor       =   16777215
            BorderStyle     =   0
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   24
            RowDividerStyle =   5
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   1
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   3
            Left            =   7680
            TabIndex        =   25
            Top             =   6120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Añadir"
            Height          =   420
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   7800
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   3
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   2415
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   2
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   5520
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   1
            Left            =   4800
            TabIndex        =   24
            Top             =   6120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Eliminar"
            Height          =   420
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   7800
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nueva Venta"
            Height          =   420
            Index           =   0
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   9720
            Width           =   3015
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   2160
            TabIndex        =   23
            Top             =   6120
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   1
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   2
            Left            =   2160
            MaxLength       =   240
            TabIndex        =   26
            Top             =   6720
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   3720
            TabIndex        =   52
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   6480
            TabIndex        =   50
            Top             =   6120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   480
            TabIndex        =   49
            Top             =   6720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   480
            TabIndex        =   37
            Top             =   9840
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   3120
            TabIndex        =   36
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PU"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   35
            Top             =   6120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. de Kg"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   720
            TabIndex        =   34
            Top             =   6120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Producto"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   840
            TabIndex        =   18
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar de venta"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4335
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   5280
         Width           =   5000
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   4
            Left            =   2520
            MaxLength       =   13
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   3
            Left            =   2520
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   2160
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   2
            Left            =   2520
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   495
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   3480
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   1
            Left            =   2520
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   0
            Left            =   2520
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   2200
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RFC"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   51
            Top             =   2640
            Width           =   2205
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código Postal"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   2205
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   2205
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ciudad"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   2205
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   2145
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4335
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   5000
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1455
            Left            =   360
            TabIndex        =   43
            Top             =   1800
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2566
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            BorderStyle     =   0
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   24
            RowDividerStyle =   5
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   1
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1800
            TabIndex        =   41
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleForeColor=   0
            CalendarTrailingForeColor=   16777215
            Format          =   117964801
            CurrentDate     =   44040
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   1800
            MaxLength       =   240
            TabIndex        =   42
            Text            =   "Text3"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   375
            Index           =   1
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   3720
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cancelar venta"
            Height          =   615
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo canc."
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A partir de"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim CN As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim RS1 As New ADODB.Recordset
    
    Dim vParty_id As Double
    Dim vParty_name As String
    Dim vAddress1 As String
    Dim vCity As String
    Dim vState As String
    Dim vPostal_code As String
    Dim vRfc As String
    
    Dim vSalesRep_id As String
    
    Dim vHeader_id As Double
    Dim vOrg_id As Integer
    Dim vOrder_type_id As Integer
    Dim vOrder_Number As Integer
    Dim vPrice_list_id As Double
    Dim vPayment_term_id As Integer
    Dim vSold_from_org_id As Integer
    Dim vSold_to_org_id As Integer
    Dim vShip_from_org_id As Integer
    Dim vShip_to_org_id As Integer
    Dim vInvoice_to_org_id As Integer
    Dim vCustomer_id As Double
    Dim vCreation_date As Date
    Dim vLine_type_id As Integer
    Dim vLine_number As Integer
    Dim vOrdered_item As String
    Dim vOrder_quantity_uom As String
    Dim vOrdered_quantity As String
    Dim vCust_po_number As String
    Dim vInventory_item_id As Double
    Dim vUnit_list_price As String
    Dim vTransaction_reference As String
    
    Dim vMotivo As String
            
    Dim Str As String
    Dim ArrStr() As String
            
    Sub ValidarDirectorios()
    
    On Error GoTo ErrorDirectorio
    
    i = GetAttr(App.Path & "\ERP")
    
    Exit Sub
ErrorDirectorio:
        If Err.Number = 53 Then
            MkDir App.Path & "\ERP"
        End If
    End Sub

Private Sub Form_Load()
    With Form1
        .BackColor = &HFFFFFF
        .WindowState = 2
        
        With Command1(0)
            .BackColor = &HFFFFFF
            .Caption = "Clientes"
        End With
        
        With Command1(1)
            .BackColor = &HFFFFFF
            .Caption = "Ventas"
        End With
        
        With Command1(2)
            .BackColor = &HFFFFFF
            .Caption = "Historial de Ventas"
        End With
        
        With Command1(3)
            .BackColor = &HFFFFFF
            .Caption = "Salir"
        End With
        
        With FramePrincipal
            .BackColor = &HFFFFFF
            .Caption = ""
        End With
        
        With Frame1(0)
            .BackColor = &HFFFFFF
            .Visible = False
            .Caption = "Clientes"
        End With
        
        With Frame1(1)
            .BackColor = &HFFFFFF
            .Visible = False
            .Caption = "Ventas"
        End With
        
        With Frame1(2)
            .BackColor = &HFFFFFF
            .Visible = False
            .Caption = "Historial de ventas"
        End With
        
        With Image1
            .Visible = True
            .Picture = LoadPicture(App.Path & "\Imagenes\Inicio.jpg")
        End With
    End With
End Sub

Private Sub Form_Resize()
    With Form1
        For i = 0 To 3
            With Command1(i)
                .Height = Round(Form1.Height / 5)
                .Width = Round(Form1.Width / 3)
                .Left = Round(Form1.Width / 27)
                .FontSize = Round(Form1.Height / 250)
            End With
        Next i
        
        With Command1(0)
            .Top = Round(Form1.Height / 25)
        End With
        
        With Command1(1)
            .Top = Command1(0).Top + Command1(0).Height + Round(Form1.Height / 25)
        End With
        
        With Command1(2)
            .Top = Command1(1).Top + Command1(1).Height + Round(Form1.Height / 25)
        End With
        
        With Command1(3)
            .Top = Command1(2).Top + Command1(2).Height + Round(Form1.Height / 25)
        End With
        
        With FramePrincipal
            .Height = Command1(2).Top + Command1(2).Height * 2
            .Width = Round(Form1.Width / 1.8)
            .Left = Round(Form1.Width / 2.5)
            .Top = Round(Form1.Height / 5 / 5)
            .FontSize = Round(Form1.Height / 250)
        End With
        
        For i = 0 To 2
            With Frame1(i)
                .Height = FramePrincipal.Height - 480
                .Width = FramePrincipal.Width - 480
                .Left = 240
                .Top = 240
                .FontSize = Round(Form1.Height / 500)
            End With
        Next i
        
        With Image1
            .Height = FramePrincipal.Height - 480
            .Width = FramePrincipal.Width - 480
            .Left = 240
            .Top = 240
        End With
        
        'CLIENTES
        
        For i = 0 To 5
            With Label1(i)
                .Width = Round(Frame1(0).Width / 5)
                .Height = Round(.Width / 4)
                .Left = 240
                .FontSize = Round(.Height / 40)
            End With
        Next i
        
        With Label1(0)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label1(1)
            .Top = Round(.Height * 3)
        End With
        
        With Label1(2)
            .Top = Round(.Height * 4.5)
        End With
        
        With Label1(3)
            .Top = Round(.Height * 6)
        End With
        
        With Label1(4)
            .Top = Round(.Height * 7.5)
        End With
        
        With Label1(5)
            .Top = Round(.Height * 9)
        End With
        
        With Combo1
            .Width = Round(Frame1(0).Width / 1.4)
            .Left = Round(Frame1(0).Width / 5) + 480
            .Top = Label1(3).Top
            .FontSize = Label1(3).FontSize
        End With
        
        For i = 0 To 4
            With Text1(i)
                .Width = Combo1.Width
                .Height = Combo1.Height
                .Left = Combo1.Left
                .FontSize = Combo1.FontSize
            End With
        Next i

        With Text1(0)
            .Top = Label1(0).Top
        End With
        
        With Text1(1)
            .Top = Label1(1).Top
        End With
        
        With Text1(2)
            .Top = Label1(2).Top
        End With
        
        With Text1(3)
            .Top = Label1(4).Top
        End With
        
        With Text1(4)
            .Top = Label1(5).Top
        End With
        
        With Command2
            .Width = Frame1(0).Width / 2
            .Height = Label1(5).Height * 1.5
            .Top = Label1(5).Top + Label1(5).Height * 2
            .Left = Frame1(0).Width / 4
            .FontSize = Label1(4).FontSize
        End With
        
        'VENTAS
    
        For i = 0 To 3
            With Label2(i)
                .Width = Round(Frame1(1).Width / 5)
                .Height = Round(.Width / 4)
                .Left = 240
                .FontSize = Round(.Height / 40)
            End With
        Next i
        
        With Label2(0)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label2(1)
            .Top = Round(.Height * 3)
        End With
        
        With Label2(2)
            .Top = Round(.Height * 4.5)
        End With
        
        With Label2(3)
            .Top = Round(.Height * 6)
        End With
        
        With Label2(4)
            .Width = Round(Frame1(1).Width / 6)
            .Height = Label2(3).Height
            .Left = Frame1(1).Width / 3.1
            .FontSize = Label2(3).FontSize
            .Top = Label2(3).Top
        End With
        
        With Label2(5)
            .Width = Label2(4).Width
            .Height = Label2(4).Height
            .Left = Frame1(1).Width / 2
            .FontSize = Label2(4).FontSize
            .Top = Label2(0).Top
        End With
        
        With Label2(7)
            .Left = 240
            .Width = Round(Frame1(1).Width / 5)
            .Height = Round(.Width / 4)
            .FontSize = Label2(4).FontSize
            .Top = .Height * 7.5
        End With
        
        With Label2(8)
            .Left = Frame1(1).Width / 1.6
            .Width = Label2(4).Width
            .Height = Label2(4).Height
            .FontSize = Label2(4).FontSize
            .Top = Label2(4).Top
        End With
        
        With Label2(9)
            .Left = Label2(5).Left
            .Width = Label2(5).Width
            .Height = Label2(4).Height
            .FontSize = Label2(4).FontSize
            .Top = Label2(7).Top
        End With
        
        For i = 0 To 2
            With Combo2(i)
                .Width = Round(Frame1(1).Width / 1.4)
                .Left = Round(Frame1(1).Width / 5) + 480
                .Top = Label2(i).Top
                .FontSize = Label2(i).FontSize
            End With
        Next i
        
        With Combo2(0)
            .Width = ((Frame1(1).Width / 2) - Label2(3).Width) - 240
        End With
        
        With Combo2(3)
            .Left = Label2(5).Left + Label2(5).Width + 240
            .Width = Frame1(1).Width - .Left - (Frame1(1).Width - Combo2(1).Width - Combo2(1).Left)
            .Top = Label2(5).Top
            .FontSize = Combo2(0).FontSize
        End With
        
        With Combo2(4)
            .Left = Label2(9).Left + Label2(9).Width + 240
            .Width = Frame1(1).Width - .Left - (Frame1(1).Width - Combo2(1).Width - Combo2(1).Left)
            .Top = Label2(9).Top
            .FontSize = Combo2(0).FontSize
        End With
        
        With Text2(0)
            .Width = Combo2(3).Width / 1.5
            .Height = Combo2(0).Height
            .Left = Combo2(0).Left
            .Top = Label2(3).Top
            .FontSize = Combo2(0).FontSize
        End With
        
        With Text2(1)
            .Left = Label2(4).Left + Label2(4).Width + 240
            .Width = Combo2(3).Width / 2
            .Height = Combo2(0).Height
            .Top = Label2(4).Top
            .FontSize = Combo2(0).FontSize
        End With
        
        With Text2(2)
            .Width = Combo2(0).Width
            .Height = Combo2(1).Height
            .Left = Combo2(1).Left
            .Top = Label2(7).Top
            .FontSize = Combo2(1).FontSize
        End With
        
        With Text2(3)
            .Left = Label2(8).Left + Label2(8).Width + 240
            .Width = Text2(1).Width
            .Height = Text2(1).Height
            .Top = Text2(1).Top
            .FontSize = Text2(1).FontSize
        End With
        
        With Command3(1)
            .Width = Frame1(1).Width / 4
            .Height = Label2(2).Height
            .Left = Frame1(1).Width - Combo2(1).Width - Combo2(1).Left
            .Top = Label2(0).Height * 9
            .FontSize = Combo2(0).FontSize
        End With
        
        With Command3(2)
            .Width = Command3(1).Width
            .Height = Command3(1).Height
            .Left = Command3(1).Left + Command3(1).Width + 240
            .Top = Command3(1).Top
            .FontSize = Command3(1).FontSize
        End With
        
        With DataGrid1
            .Left = Command3(1).Left
            .Width = Frame1(1).Width - .Left * 2
            .Height = Text2(0).Height * 4
            .Top = Label2(0).Height * 10.5
            Set .Font = Combo2(0).Font
        End With
        
        With Label2(6)
            .Left = DataGrid1.Left
            .Width = DataGrid1.Width
            .Height = Label2(0).Height
            .Top = DataGrid1.Top + DataGrid1.Height + Label2(0).Height * 0.5
            .FontSize = Combo2(0).FontSize
        End With
        
        With Command3(0)
            .Width = Frame1(1).Width / 2
            .Height = Label2(2).Height * 1.5
            .Top = Label2(6).Top + Label2(6).Height + Label2(6).Height * 0.5
            .Left = Frame1(1).Width / 4
            .FontSize = Label2(2).FontSize
        End With
        
        'HISTORIAL DE VENTAS
        
        For i = 0 To 2
            With Label3(i)
                .Width = Round(Frame1(2).Width / 5)
                .Height = Round(.Width / 4)
                .Left = 240
                .FontSize = Round(.Height / 40)
            End With
        Next i
        
        With Label3(0)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label3(1)
            .Top = Round(.Height * 3)
        End With
        
        With Label3(2)
            .Top = Round(.Height * 4.5)
        End With
        
        With Combo3
            .Width = Round(Frame1(2).Width / 1.4)
            .Left = Round(Frame1(2).Width / 5) + 480
            .Top = Label3(0).Top
            .FontSize = Label3(0).FontSize
        End With
        
        With DTPicker1
            .Width = Combo3.Width
            .Height = Combo3.Height
            .Left = Combo3.Left
            .Top = Label3(1).Top
            Set .Font = Combo3.Font
        End With
        
        With Text3
            .Width = Combo3.Width
            .Height = Combo3.Height
            .Left = Combo3.Left
            .Top = Label3(2).Top
            .FontSize = Combo3.FontSize
        End With
        
        With DataGrid2
            .Left = Frame1(2).Width - Combo3.Width - Combo3.Left
            .Width = Frame1(2).Width - .Left * 2
            .Height = Text3.Height * 6
            .Top = Label3(0).Height * 6
            Set .Font = Combo3.Font
        End With
        
        For i = 0 To 1
            With Command4(i)
                .Width = Frame1(2).Width / 2.5
                .Height = Label3(0).Height * 1.5
                .Top = DataGrid2.Top + DataGrid2.Height + (Label3(0).Height * 1.5)
                .FontSize = Label3(0).FontSize
            End With
        Next i
        
        With Command4(0)
            .Left = DataGrid2.Left
        End With
        
        With Command4(1)
            .Left = Frame1(2).Width - (Frame1(2).Width - Combo3.Width - Combo3.Left) - .Width
        End With
    End With
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        
        Case 0
            If Frame1(0).Visible = True Then
                MsgBox "La pantalla ya está abierta", vbOKOnly, "Información"
                Exit Sub
            End If
            
            If Frame1(1).Visible = True Or Frame1(2).Visible = True Then
                vbq = MsgBox("¿Desea cerrar la pantalla abierta para abrir la pantalla de Clientes?", vbQuestion + vbYesNo, "Advertencia")
                
                If vbq = vbNo Then
                    Exit Sub
                End If
            End If
            
            With DataGrid1
                Set .DataSource = Nothing
            End With
            
            With DataGrid2
                Set .DataSource = Nothing
            End With
            
            With RS
                If .State = 1 Then .Close
            End With
            
            With CN
                If .State = 1 Then .Close
            End With
            
            Set RS = Nothing
            
            Set CN = Nothing

                
            With Frame1(0)
                .Visible = True
            End With
                
            With Frame1(1)
                .Visible = False
            End With
            
            With Frame1(2)
                .Visible = False
            End With
            
            With Image1
                .Visible = False
            End With
            
            With Combo1
                .Clear
            End With
            
            For i = 0 To 4
                With Text1(i)
                    .Text = ""
                End With
            Next i
            
            With CN
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
            End With
                
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from HZ_STATES_ALL Order by 2", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo1.AddItem .Fields(1).Value
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With CN
                If .State = 1 Then .Close
            End With
        Case 1
            If Frame1(1).Visible = True Then
                MsgBox "La pantalla ya está abierta", vbOKOnly, "Información"
                Exit Sub
            End If
            
            If Frame1(0).Visible = True Or Frame1(2).Visible = True Then
                vbq = MsgBox("¿Desea cerrar la pantalla abierta para abrir la pantalla Ventas?", vbQuestion + vbYesNo, "Advertencia")
                
                If vbq = vbNo Then
                    Exit Sub
                End If
            End If
            
            With DataGrid1
                Set .DataSource = Nothing
            End With
            
            With DataGrid2
                Set .DataSource = Nothing
            End With
            
            With RS
                If .State = 1 Then .Close
            End With
            
            With CN
                If .State = 1 Then .Close
            End With
            
            Set RS = Nothing
            
            Set CN = Nothing
                
            With Frame1(0)
                .Visible = False
            End With
                
            With Frame1(1)
                .Visible = True
            End With
            
            With Frame1(2)
                .Visible = False
            End With
            
            With Image1
                .Visible = False
            End With
            
            For i = 0 To 4
                With Combo2(i)
                    .Clear
                End With
            Next i
            
            For i = 0 To 2
                With Text2(i)
                    .Text = ""
                End With
            Next i
            
            With Combo2(3)
                .AddItem "CONTADO"
                .AddItem "CRÉDITO"
            End With
            
            With CN
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select count(*) + 1 from OE_ORDER_HEADERS_ALL", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    Frame1(1).Caption = "Nueva Venta            Folio [DEP-" & .Fields(0).Value & "]"
                End If
                
                .Close
            End With
                
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from HR_ALL_ORGANIZATION_UNITS Order by NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo2(0).AddItem .Fields(1).Value & " [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from HZ_PARTIES Order by PARTY_NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo2(1).AddItem .Fields(2).Value & " [" & .Fields(1).Value & "] [" & .Fields(0).Value & "] [" & .Fields(7).Value & "] [" & .Fields(8).Value & "] [" & .Fields(9).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With Combo2(2)
                .AddItem ""
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from MTL_SYSTEM_ITEMS_B Where ENABLED_FLAG = 'Y' Order by DESCRIPTION", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo2(2).AddItem .Fields(2).Value & " [" & .Fields(1).Value & "] [" & .Fields(0).Value & "] [" & .Fields(4).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from JTF_RS_SALESREP Order by NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo2(4).AddItem .Fields(1).Value & " [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With DataGrid1
                Set .DataSource = Nothing
            End With
            
            With Label2(6)
                .Caption = "TOTAL $ 0.00"
            End With
            
            vLine_number = 1
            
            With Combo2(0)
                .Enabled = True
            End With
            
            With Combo2(1)
                .Enabled = True
            End With
            
            With Combo2(3)
                .Enabled = True
            End With
            
            With Text2(2)
                .Enabled = True
            End With
        Case 2
            If Frame1(2).Visible = True Then
                MsgBox "La pantalla ya está abierta", vbOKOnly, "Información"
                Exit Sub
            End If
            
            If Frame1(0).Visible = True Or Frame1(1).Visible = True Then
                vbq = MsgBox("¿Desea cerrar la pantalla abierta para abrir la pantalla de historial de ventas?", vbQuestion + vbYesNo, "Advertencia")
                
                If vbq = vbNo Then
                    Exit Sub
                End If
            End If
            
            With DataGrid1
                Set .DataSource = Nothing
            End With
            
            With DataGrid2
                Set .DataSource = Nothing
            End With
            
            With RS
                If .State = 1 Then .Close
            End With
            
            With CN
                If .State = 1 Then .Close
            End With
            
            Set RS = Nothing
            
            Set CN = Nothing
                
            With Frame1(0)
                .Visible = False
            End With
                
            With Frame1(1)
                .Visible = False
            End With
            
            With Frame1(2)
                .Visible = True
            End With
            
            With Image1
                .Visible = False
            End With
            
            With Combo3
                .Clear
                .AddItem ""
            End With
            
            With Text3
                .Text = ""
                .Enabled = False
            End With
            
            With DTPicker1
                .Value = Date
            End With
            
            With CN
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
            End With
                
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from HZ_PARTIES Order by PARTY_NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo3.AddItem .Fields(2).Value
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select T1.* from OE_ORDER_HEADERS_ALL_V T1 Where NOT EXISTS (Select * from OE_ORDER_CANCEL_ALL Where HEADER_ID = T1.HEADER_ID) Order by 1", CN, adOpenStatic, adLockOptimistic
                .Filter = "CREATION_DATE >= '" & DTPicker1.Value & "'"
                .Requery
            End With
            
            With DataGrid2
                Set .DataSource = RS
                
                With .Columns(0)
                    .Visible = False
                    .Width = DataGrid2.Width / 4.5
                End With
                
                With .Columns(1)
                    .Caption = "FOLIO"
                    .Width = DataGrid2.Width / 6
                End With
                
                With .Columns(2)
                    .Caption = "CLIENTE"
                    .Width = DataGrid2.Width / 2.5
                End With
                
                With .Columns(3)
                    .Caption = "FECHA"
                    .Width = DataGrid2.Width / 6
                End With
                
                With .Columns(4)
                    .Caption = "TOTAL"
                    .Width = DataGrid2.Width / 5
                    .Alignment = dbgRight
                End With
                
                With .Columns(5)
                    .Visible = False
                End With
                
                With .Columns(6)
                    .Visible = False
                End With
            End With
            
            With Command4(1)
                .Enabled = False
            End With
        Case 3
            vbq = MsgBox("¿Desea cerrar el programa?", vbQuestion + vbYesNo, "Advertencia")
                    
            If vbq = vbYes Then
                With DataGrid1
                    Set .DataSource = Nothing
                End With
                
                With DataGrid2
                    Set .DataSource = Nothing
                End With
                
                With RS
                    If .State = 1 Then .Close
                End With
                
                With CN
                    If .State = 1 Then .Close
                End With
                
                Set RS = Nothing
                
                Set CN = Nothing
            
                ValidarDirectorios
                
                With CN
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
                End With
                    
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select * from HZ_PARTIES Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\HP_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                        
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|" & .Fields(10).Value & "|"
                            
                            .Fields(11).Value = 1
                            .Update
                            .MoveNext
                        Wend
                        
                        Close #1
                    End If
                    
                    .Close
                End With
                
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select * from OE_ORDER_HEADERS_ALL Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\OOHA_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                       
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|" & .Fields(10).Value & "|" & .Fields(11).Value & "|" & .Fields(12).Value & "|" & .Fields(13).Value & "|" & .Fields(14).Value & "|" & .Fields(15).Value & "|" & .Fields(16).Value & "|"
                            
                            .Fields(17).Value = 1
                            .Update
                            .MoveNext
                        Wend
                        
                        Close #1
                    End If
                    
                    .Close
                End With
                
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select * from OE_ORDER_LINES_ALL Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\OOLA_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                       
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|" & .Fields(10).Value & "|" & .Fields(11).Value & "|" & .Fields(12).Value & "|" & .Fields(13).Value & "|" & .Fields(14).Value & "|" & .Fields(15).Value & "|" & .Fields(16).Value & "|"
                            
                            .Fields(17).Value = 1
                            .Update
                            .MoveNext
                        Wend
                        
                        Close #1
                    End If
                    
                    .Close
                End With
                
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select * from OE_ORDER_CANCEL_ALL Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\OOCA_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                       
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|"
                            
                            .Fields(2).Value = 1
                            .Update
                            .MoveNext
                        Wend
                        
                        Close #1
                    End If
                    
                    .Close
                End With
                
                With DataGrid1
                    Set .DataSource = Nothing
                End With
                
                With DataGrid2
                    Set .DataSource = Nothing
                End With
                
                With RS
                    If .State = 1 Then .Close
                End With
                
                With CN
                    If .State = 1 Then .Close
                End With
                
                Set RS = Nothing
                
                Set CN = Nothing
            
                Unload Me
            End If
    End Select
End Sub

Private Sub Command2_Click()
    With Text1(0)
        If .Text = "" Then
            MsgBox "Por introduzca el nombre del cliente", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text1(1)
        If .Text = "" Then
            MsgBox "Por introduzca la dirección", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text1(2)
        If .Text = "" Then
            MsgBox "Por introduzca la ciudad", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Combo1
        If .Text = "" Then
            MsgBox "Por favor seleccione el estado", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text1(3)
        If .Text = "" Then
            MsgBox "Por introduzca el código postal", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text1(4)
        If .Text = "" Then
            MsgBox "Por introduzca el RFC", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        Else
            If Len(.Text) < 12 Then
                MsgBox "RFC Inválido, longitud minima 12 caracteres", vbOKOnly, "Advertencia"
                
                .SetFocus
                
                Exit Sub
            Else
                If .Text = "XAXX010101000" Or .Text = "XEXX010101000" Then
                    MsgBox "RFC no permitido", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                Else
                    Dim vExiste As Integer
                    
                    With CN
                        If .State = 1 Then .Close
                        .CursorLocation = adUseClient
                        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
                    End With
                    
                    With RS
                        If .State = 1 Then .Close
                        .CursorLocation = adUseClient
                        .Open "Select count (*) from HZ_PARTIES where rfc = '" & Text1(4).Text & "'", CN, adOpenStatic, adLockOptimistic
                        .Requery
                        
                        vExiste = .Fields(0).Value
                        
                        .Close
                    End With
                    
                    With CN
                        If .State = 1 Then .Close
                    End With
                
                    If vExiste > 0 Then
                        MsgBox "El RFC esta duplicado", vbOKOnly, "Advertencia"
                        
                        .SetFocus
                        
                        Exit Sub
                    End If
                End If
            End If
        End If
    End With
    
    With CN
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
    End With
    
    With RS
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Select count (*) + 1 from HZ_PARTIES", CN, adOpenStatic, adLockOptimistic
        .Requery
        
        vParty_id = .Fields(0).Value
        
        .Close
    End With
    
    vParty_name = Text1(0).Text
    vAddress1 = Text1(1).Text
    vCity = Text1(2).Text
    vState = Combo1.Text
    vPostal_code = Text1(3).Text
    vRfc = Text1(4).Text
     
    With RS
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Select * from HZ_PARTIES", CN, adOpenStatic, adLockOptimistic
        .AddNew
            .Fields(0) = vParty_id
            .Fields(2) = vParty_name
            .Fields(3) = vAddress1
            .Fields(4) = vCity
            .Fields(5) = vState
            .Fields(6) = vPostal_code
            .Fields(10) = vRfc
        .Update
        .Requery
        .Close
    End With
    
    With CN
        If .State = 1 Then .Close
    End With
    
    MsgBox "Cliente guardado con éxito", vbOKOnly, "Terminado"
    
    With Frame1(0)
        .Visible = False
    End With
    
    With Image1
        .Visible = True
    End With
End Sub

Private Sub Command3_Click(Index As Integer)
    Select Case Index
        Case 1
            With Combo2(0)
                If .Text = "" Then
                    MsgBox "Por favor seleccione el lugar de la venta", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Combo2(1)
                If .Text = "" Then
                    MsgBox "Por favor seleccione el cliente", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Combo2(1)
                If .Text = "" Then
                    MsgBox "Por favor seleccione el cliente", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Combo2(3)
                If .Text = "" Then
                    MsgBox "Por favor seleccione el tipo de venta", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Combo2(4)
                If .Text = "" Then
                    MsgBox "Por favor seleccione un vendedor", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Text2(0)
                If .Text = "" Then
                    MsgBox "Por introduzca la cantidad de kilogramos", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
                
                If Val(.Text) <= 0 Then
                    MsgBox "La cantidad de kilogramos ingresada no es válida, por favor corrija la información", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Text2(1)
                If .Text = "" Then
                    MsgBox "Por introduzca el precio", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
                
                If Val(.Text) < 0 Then
                    MsgBox "El precio ingresado no es válido, por favor corrija la información", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Text2(3)
                If .Text = "" Then
                    MsgBox "Por introduzca el descuento", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
                
                If Val(.Text) < 0 Then
                    MsgBox "El descuento ingresado no es válido, por favor corrija la información", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
                
                If Val(.Text) = 0 Then
                    vbq = MsgBox("¿El descuento es de 0 pesos, desea continuar?", vbQuestion + vbYesNo, "Advertencia")
                
                    If vbq = vbNo Then
                        Exit Sub
                    End If
                End If
            End With
            
            vOrg_id = 490
            vOrder_type_id = 1133 'PRODUCCION
            'vOrder_type_id = 1113 'DESARROLLO
            vPrice_list_id = 363123
            
            If Combo2(3).Text = "CONTADO" Then
                vPayment_term_id = 1000
            Else
                vPayment_term_id = 1003
            End If
            
            vSold_from_org_id = 490
            
            Str = Combo2(1).Text
            ArrStr() = Split(Str, "[")
            
            vSold_to_org_id = Replace(ArrStr(4), "]", "")
            vShip_to_org_id = Replace(ArrStr(5), "]", "")
            vInvoice_to_org_id = Replace(ArrStr(4), "]", "")
            vCustomer_id = Replace(ArrStr(3), "]", "")
            vParty_id = Replace(ArrStr(2), "]", "")
            
            Str = Combo2(0).Text
            ArrStr() = Split(Str, "[")
            
            vShip_from_org_id = Replace(ArrStr(1), "]", "")
            
            vCreation_date = Date
            
            'vLine_type_id = 1112 'DESARROLLO
            vLine_type_id = 1132 'PRODUCCION
            
            vOrder_quantity_uom = "KG"
            vOrdered_quantity = Text2(0).Text
            
            Str = Combo2(2).Text
            ArrStr() = Split(Str, "[")
            
            vOrdered_item = Trim(Replace(ArrStr(1), "]", ""))
            vInventory_item_id = Replace(ArrStr(2), "]", "")
            vUnit_list_price = Replace(Val(Text2(1).Text) - Val(Text2(3).Text), ",", ".")
            vTransaction_reference = Text2(2).Text
            
            Str = Combo2(4).Text
            ArrStr() = Split(Str, "[")
            
            vSalesRep_id = Replace(ArrStr(1), "]", "")
             
            If vLine_number = 1 Then
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select count(*) + 1 from OE_ORDER_HEADERS_ALL", CN, adOpenStatic, adLockOptimistic
                
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        
                        vHeader_id = .Fields(0).Value
                        vOrder_Number = .Fields(0).Value
                        vCust_po_number = "DEP-" & .Fields(0).Value
                    End If
                    
                    .Close
                End With
            
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select * from OE_ORDER_HEADERS_ALL", CN, adOpenStatic, adLockOptimistic
                    .AddNew
                        .Fields(0) = vHeader_id
                        .Fields(1) = vOrg_id
                        .Fields(2) = vOrder_type_id
                        .Fields(3) = vOrder_Number
                        .Fields(4) = vPrice_list_id
                        .Fields(5) = vPayment_term_id
                        .Fields(6) = vSold_from_org_id
                        .Fields(7) = vSold_to_org_id
                        .Fields(8) = vShip_from_org_id
                        .Fields(9) = vShip_to_org_id
                        .Fields(10) = vInvoice_to_org_id
                        .Fields(11) = vCustomer_id
                        .Fields(12) = vParty_id
                        .Fields(13) = vCreation_date
                        .Fields(14) = vCust_po_number
                        .Fields(15) = vTransaction_reference
                        .Fields(16) = vSalesRep_id
                    .Update
                    .Requery
                    .Close
                End With
            End If
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from OE_ORDER_LINES_ALL", CN, adOpenStatic, adLockOptimistic
                .AddNew
                    .Fields(1) = vOrg_id
                    .Fields(2) = vHeader_id
                    .Fields(3) = vLine_type_id
                    .Fields(4) = vLine_number
                    .Fields(5) = vOrdered_item
                    .Fields(6) = vOrder_quantity_uom
                    .Fields(7) = vOrdered_quantity
                    .Fields(8) = vShip_from_org_id
                    .Fields(9) = vSold_from_org_id
                    .Fields(10) = vSold_to_org_id
                    .Fields(11) = vCust_po_number
                    .Fields(12) = vInventory_item_id
                    .Fields(13) = vPrice_list_id
                    .Fields(14) = vUnit_list_price
                    .Fields(15) = vShip_to_org_id
                    .Fields(16) = vInvoice_to_org_id
                .Update
                .Requery
                .Close
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from OE_ORDER_HEADERS_ALL_V", CN, adOpenStatic, adLockOptimistic
                .Filter = "HEADER_ID =" & vHeader_id
                .Requery
                
                Label2(6).Caption = "TOTAL $ " & .Fields(4).Value
            End With
            
            
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from OE_ORDER_LINES_ALL_V", CN, adOpenStatic, adLockOptimistic
                .Filter = "HEADER_ID =" & vHeader_id
                .Requery
            End With
            
            With DataGrid1
                Set .DataSource = RS
                
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                
                For i = 0 To 6
                    .Columns(i).Width = .Width / 5.5
                Next i
                
                .Columns(2).Width = .Width / 5.2
                
                .Columns(2).Caption = "CÓDIGO"
                
                With .Columns(3)
                    .Caption = "PRECIO"
                    .Alignment = dbgRight
                End With
                
                With .Columns(4)
                    .Caption = "CANTIDAD"
                    .Alignment = dbgRight
                End With
                
                .Columns(5).Caption = "UDM"
                
                With .Columns(6)
                    .Caption = "SUBTOTAL"
                    .Alignment = dbgRight
                End With
            End With
            
            vLine_number = vLine_number + 1
            
            With Combo2(0)
                .Enabled = False
            End With
            
            With Combo2(1)
                .Enabled = False
            End With
            
            With Combo2(3)
                .Enabled = False
            End With
            
            With Combo2(4)
                .Enabled = False
            End With
            
            With Text2(2)
                .Enabled = False
            End With
            
            For i = 0 To 1
                With Text2(i)
                    .Text = ""
                End With
            Next i
            
            With Combo2(2)
                .ListIndex = 0
                .SetFocus
            End With
        Case 2
            With RS
                .Delete
                .Requery
            End With
            
            With DataGrid1
                Set .DataSource = RS
                
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                
                For i = 0 To 6
                    .Columns(i).Width = .Width / 5.5
                Next i
                
                .Columns(2).Width = .Width / 5.2
                
                .Columns(2).Caption = "CÓDIGO"
                
                With .Columns(3)
                    .Caption = "PRECIO"
                    .Alignment = dbgRight
                End With
                
                With .Columns(4)
                    .Caption = "CANTIDAD"
                    .Alignment = dbgRight
                End With
                
                .Columns(5).Caption = "UDM"
                
                With .Columns(6)
                    .Caption = "SUBTOTAL"
                    .Alignment = dbgRight
                End With
            End With
        Case 0
            MsgBox "Venta guardada con éxito", vbOKOnly, "Terminado"
            
            With DataGrid1
                Set .DataSource = Nothing
            End With
            
            With RS
                If .State = 1 Then .Close
            End With
            
            Set RS = Nothing
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from OE_TICKET_ALL Where CUST_PO_NUMBER = '" & vCust_po_number & "'", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    With DataReport1
                        Set .DataSource = RS
                        
                        With .Sections("Sección4")
                            .Controls("Etiqueta6").Caption = RS.Fields(1).Value
                            .Controls("Etiqueta11").Caption = RS.Fields(0).Value
                            .Controls("Etiqueta12").Caption = RS.Fields(2).Value
                            
                            If RS.Fields(8).Value = 1000 Then
                                .Controls("Label1").Caption = "NOTA DE RECIBO"
                            Else
                                .Controls("Label1").Caption = "NOTA DE VENTA"
                            End If
                        End With
                        
                        With .Sections("Sección1")
                            .Controls("Texto1").DataField = "DESCRIPTION"
                            .Controls("Texto3").DataField = "UNIT_LIST_PRICE"
                            .Controls("Texto5").DataField = "ORDERED_QUANTITY"
                            .Controls("Texto6").DataField = "SUBTOTAL"
                        End With
                        
                        With .Sections("Sección5")
                            If RS.Fields(8).Value = 1000 Then
                                .Controls("Etiqueta8").Visible = False
                                .Controls("Etiqueta8").Height = 0
                                .Controls("Etiqueta18").Top = 1700
                                .Controls("Etiqueta17").Top = 2268
                            Else
                                .Controls("Etiqueta8").Visible = True
                                .Controls("Etiqueta8").Height = 5670
                                .Controls("Etiqueta18").Top = 7370
                                .Controls("Etiqueta17").Top = 7938
                            End If
                            
                            .Controls("Etiqueta9").Caption = "TOTAL $" & RS.Fields(7).Value
                        End With
                        .Show 1
                    End With
                End If
                
                .Close
            End With
            
            With CN
                If .State = 1 Then .Close
            End With
            
            With Frame1(1)
                .Visible = False
            End With
            
            With Image1
                .Visible = True
            End With
    End Select
End Sub

Private Sub Combo2_Click(Index As Integer)
    Select Case Index
        Case 2
            Str = Combo2(2).Text
            ArrStr() = Split(Str, "[")
            
            If Combo2(2).Text <> "" Then
                Text2(1).Text = Replace(ArrStr(3), "]", "")
                Text2(3).Text = "0"
            Else
                Text2(1).Text = ""
                Text2(3).Text = ""
            End If
    End Select

End Sub

Private Sub Combo3_Click()
    With Combo3
        If .Text = "" Then
            With RS
                .Filter = "CREATION_DATE >= '" & DTPicker1.Value & "'"
                .Requery
            End With
        Else
            With RS
                .Filter = "PARTY_NAME = '" & Combo3.Text & "' AND CREATION_DATE >= '" & DTPicker1.Value & "'"
                .Requery
            End With
        End If
    End With
    
    With DataGrid2
        Set .DataSource = RS
                
        With .Columns(0)
            .Visible = False
            .Width = DataGrid2.Width / 4.5
        End With
                
        With .Columns(1)
            .Caption = "FOLIO"
            .Width = DataGrid2.Width / 6
        End With
                
        With .Columns(2)
            .Caption = "CLIENTE"
            .Width = DataGrid2.Width / 2.5
        End With
                
        With .Columns(3)
            .Caption = "FECHA"
            .Width = DataGrid2.Width / 6
        End With
                
        With .Columns(4)
            .Caption = "TOTAL"
            .Width = DataGrid2.Width / 5
            .Alignment = dbgRight
        End With
        
        With .Columns(5)
            .Visible = False
        End With
                
        With .Columns(6)
            .Visible = False
        End With
    End With
End Sub

Private Sub DTPicker1_Change()
    With Combo3
        If .Text = "" Then
            With RS
                .Filter = "CREATION_DATE >= '" & DTPicker1.Value & "'"
                .Requery
            End With
        Else
            With RS
                .Filter = "PARTY_NAME = '" & Combo3.Text & "' AND CREATION_DATE >= '" & DTPicker1.Value & "'"
                .Requery
            End With
        End If
    End With
    
    With DataGrid2
        Set .DataSource = RS
                
        With .Columns(0)
            .Visible = False
            .Width = DataGrid2.Width / 4.5
        End With
                
        With .Columns(1)
            .Caption = "FOLIO"
            .Width = DataGrid2.Width / 6
        End With
                
        With .Columns(2)
            .Caption = "CLIENTE"
            .Width = DataGrid2.Width / 2.5
        End With
                
        With .Columns(3)
            .Caption = "FECHA"
            .Width = DataGrid2.Width / 6
        End With
                
        With .Columns(4)
            .Caption = "TOTAL"
            .Width = DataGrid2.Width / 5
            .Alignment = dbgRight
        End With
        
        With .Columns(5)
            .Visible = False
        End With
                
        With .Columns(6)
            .Visible = False
        End With
    End With
End Sub

Private Sub Command4_Click(Index As Integer)
    Select Case Index
        Case 0
            vbq = MsgBox("¿Desea cancelar la venta?", vbQuestion + vbYesNo, "Advertencia")
                
            If vbq = vbNo Then
                With Text3
                    .Enabled = False
                End With
                
                With Command4(0)
                    .Enabled = True
                End With
                
                With Command4(1)
                    .Enabled = False
                End With
            Else
                With Command4(0)
                    .Enabled = False
                End With
                
                With Command4(1)
                    .Enabled = True
                End With
                
                With Text3
                    .Enabled = True
                    .SetFocus
                End With
            End If
        Case 1
            With Text3
                If .Text = "" Then
                    MsgBox "Por introduzca el motivo de la cancelación", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            With Command4(0)
                .Enabled = True
            End With
                
            With Command4(1)
                .Enabled = False
            End With
            
            vHeader_id = RS.Fields(0).Value
            vCust_po_number = RS.Fields(1).Value
            vMotivo = Text3.Text
            
            With RS1
                If .State = 1 Then .Close
                
                Set RS = Nothing
                
                .CursorLocation = adUseClient
                .Open "SELECT * from OE_ORDER_CANCEL_ALL", CN, adOpenStatic, adLockOptimistic
                .AddNew
                    .Fields(0) = vHeader_id
                    .Fields(1) = vMotivo
                .Update
                .Requery
                .Close
            End With
            
            MsgBox "Venta " & vCust_po_number & " cancelada con éxito", vbOKOnly, "Terminado"
            
            With Frame1(2)
                .Visible = False
            End With
            
            With Image1
                .Visible = True
            End With
    End Select
End Sub
