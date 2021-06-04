VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   ControlBox      =   0   'False
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
   ScaleHeight     =   5325
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   5480
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5000
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Añadir"
            Height          =   420
            Index           =   1
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2520
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            Height          =   420
            Index           =   2
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   420
            Index           =   1
            Left            =   3240
            TabIndex        =   7
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Eliminar"
            Height          =   420
            Index           =   2
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   420
            Index           =   0
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3840
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   420
            Index           =   0
            Left            =   1440
            TabIndex        =   6
            Top             =   1800
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   735
            Left            =   240
            TabIndex        =   10
            Top             =   2760
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   1296
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Total"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   14
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Precio"
            Height          =   375
            Index           =   4
            Left            =   2040
            TabIndex        =   13
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Cant. de Kg"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Producto"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   5000
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1800
            TabIndex        =   18
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
            Format          =   131661825
            CurrentDate     =   44040
         End
         Begin VB.TextBox Text3 
            Height          =   420
            Left            =   1800
            MaxLength       =   240
            TabIndex        =   19
            Text            =   "Text3"
            Top             =   1320
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1455
            Left            =   360
            TabIndex        =   20
            Top             =   1800
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2566
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   375
            Index           =   1
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   3720
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            Height          =   420
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   17
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
            TabIndex        =   22
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Motivo canc."
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "A partir de"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Cliente"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   21
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
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
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
    Dim vLote As String
    
    Dim vMotivo As String
            
    Dim Str As String
    Dim ArrStr() As String
            
    Sub ValidarDirectorios()
    
    On Error GoTo ErrorDirectorio
    
    I = GetAttr(App.Path & "\ERP")
    
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
        
        With Command1(1)
            .BackColor = &HC0C0C0
            .Caption = "Nueva Compra"
        End With
        
        With Command1(2)
            .BackColor = &HC0C0C0
            .Caption = "Historial de Compras"
        End With
        
        With Command1(3)
            .BackColor = &HC0C0C0
            .Caption = "Salir"
        End With
        
        With FramePrincipal
            .BackColor = &HC0C0C0
            .Caption = ""
        End With
        
        With Frame1(1)
            .BackColor = &H8000000F
            .Visible = False
            .Caption = "Nueva Compra"
        End With
        
        With Frame1(2)
            .BackColor = &H8000000F
            .Visible = False
            .Caption = "Historial de compras"
        End With
        
        With Image1
            .Visible = True
            .Picture = LoadPicture(App.Path & "\Imagenes\Inicio.jpg")
        End With
    End With
End Sub

Private Sub Form_Resize()
    With Form1
        For I = 1 To 3
            With Command1(I)
                .Height = Round(Form1.Height / 3.5)
                .Width = Round(Form1.Width / 3)
                .Left = Round(Form1.Width / 27)
                .FontSize = Round(Form1.Height / 250)
            End With
        Next I
        
        With Command1(1)
            .Top = Round(Form1.Height / 33)
        End With
        
        With Command1(2)
            .Top = Command1(1).Top + Command1(1).Height + Round(Form1.Height / 33)
        End With
        
        With Command1(3)
            .Top = Command1(2).Top + Command1(2).Height + Round(Form1.Height / 33)
        End With
        
        With FramePrincipal
            .Height = Command1(2).Top + Command1(2).Height * 2
            .Width = Round(Form1.Width / 1.8)
            .Left = Round(Form1.Width / 2.5)
            .Top = Round(Form1.Height / 5 / 5)
            .FontSize = Round(Form1.Height / 250)
        End With
        
        For I = 1 To 2
            With Frame1(I)
                .Height = FramePrincipal.Height - 480
                .Width = FramePrincipal.Width - 480
                .Left = 240
                .Top = 240
                .FontSize = Round(Form1.Height / 500)
            End With
        Next I
        
        With Image1
            .Height = FramePrincipal.Height - 480
            .Width = FramePrincipal.Width - 480
            .Left = 240
            .Top = 240
        End With
    
        For I = 2 To 3
            With Label2(I)
                .Width = Round(Frame1(1).Width / 5)
                .Height = Round(.Width / 4)
                .Left = 240
                .FontSize = Round(.Height / 32)
            End With
        Next I
        
        With Label2(2)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label2(3)
            .Top = Round(.Height * 3)
        End With
        
        With Label2(4)
            .Width = Round(Frame1(1).Width / 6)
            .Height = Label2(3).Height
            .Left = Frame1(1).Width / 2 + 240
            .FontSize = Label2(3).FontSize
            .Top = Label2(3).Top
        End With
        
        With Combo2(2)
            .Width = Round(Frame1(1).Width / 1.4)
            .Left = Round(Frame1(1).Width / 5) + 480
            .Top = Label2(2).Top
            .FontSize = Label2(2).FontSize
        End With
        
        With Text2(0)
            .Width = Combo2(2).Width / 3
            .Height = Combo2(2).Height
            .Left = Combo2(2).Left
            .Top = Label2(3).Top
            .FontSize = Combo2(2).FontSize
        End With
        
        With Text2(1)
            .Left = Combo2(2).Left + Combo2(2).Width - (Combo2(2).Width / 3)
            .Width = Combo2(2).Width / 3
            .Height = Combo2(2).Height
            .Top = Label2(4).Top
            .FontSize = Combo2(2).FontSize
        End With
        
        With Command3(1)
            .Width = Frame1(1).Width / 4
            .Height = Label2(2).Height
            .Top = Label2(2).Height * 4.5
            .FontSize = Combo2(2).FontSize
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
            .Height = Text2(0).Height * 8
            .Top = Label2(2).Height * 6
        End With
        
        With Label2(6)
            .Left = DataGrid1.Left
            .Width = DataGrid1.Width
            .Height = Label2(2).Height
            .Top = DataGrid1.Top + DataGrid1.Height + Label2(2).Height * 0.5
            .FontSize = Command3(1).FontSize
        End With
        
        With Command3(0)
            .Width = Frame1(1).Width / 2
            .Height = Label2(2).Height * 1.5
            .Top = Label2(6).Top + Label2(6).Height + Label2(6).Height * 0.5
            .Left = Frame1(1).Width / 4
            .FontSize = Label2(2).FontSize
        End With
        
        For I = 0 To 2
            With Label3(I)
                .Width = Round(Frame1(2).Width / 5)
                .Height = Round(.Width / 4)
                .Left = 240
                .FontSize = Round(.Height / 32)
            End With
        Next I
        
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
        
        For I = 0 To 1
            With Command4(I)
                .Width = Frame1(2).Width / 2.5
                .Height = Label3(0).Height * 1.5
                .Top = DataGrid2.Top + DataGrid2.Height + (Label3(0).Height * 1.5)
                .FontSize = Label3(0).FontSize
            End With
        Next I
        
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
        Case 1
            If Frame1(1).Visible = True Then
                MsgBox "La pantalla ya está abierta", vbOKOnly, "Información"
                Exit Sub
            End If
            
            If Frame1(2).Visible = True Then
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
                
            With Frame1(1)
                .Visible = True
            End With
            
            With Frame1(2)
                .Visible = False
            End With
            
            With Image1
                .Visible = False
            End With
            
            With Combo2(2)
                .Clear
            End With
            
            For I = 0 To 1
                With Text2(I)
                    .Text = ""
                End With
            Next I
            
            With CN
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD_VRN.mdb; Persist Security Info=False"
            End With
            
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select count(*) + 1 from OE_ORDER_HEADERS_ALL", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    Frame1(1).Caption = "Nueva Venta                                                Folio [VRN-" & .Fields(0).Value & "]"
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
            
            With DataGrid1
                Set .DataSource = Nothing
            End With
            
            With Label2(6)
                .Caption = "TOTAL $ 0.00"
            End With
            
            vLine_number = 1
        Case 2
            If Frame1(2).Visible = True Then
                MsgBox "La pantalla ya está abierta", vbOKOnly, "Información"
                Exit Sub
            End If
            
            If Frame1(1).Visible = True Then
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
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD_VRN.mdb; Persist Security Info=False"
            End With
                
            With RS
                If .State = 1 Then .Close
                .CursorLocation = adUseClient
                .Open "Select * from HZ_PARTIES Order by PARTY_NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo3.AddItem .Fields(2).Value '& " [" & .Fields(0).Value & "]"
                        
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
                    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD_VRN.mdb; Persist Security Info=False"
                End With
                    
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select * from HZ_PARTIES Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\HP_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                        
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|"
                            
                            .Fields(10).Value = 1
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
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|" & .Fields(10).Value & "|" & .Fields(11).Value & "|" & .Fields(12).Value & "|" & .Fields(13).Value & "|" & .Fields(14).Value & "|" & .Fields(15).Value & "|"
                            
                            .Fields(16).Value = 1
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
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|" & .Fields(10).Value & "|" & .Fields(11).Value & "|" & .Fields(12).Value & "|" & .Fields(13).Value & "|" & .Fields(14).Value & "|" & .Fields(15).Value & "|" & .Fields(16).Value & "|" & .Fields(18).Value & "|"
                            
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

Private Sub Command3_Click(Index As Integer)
    Select Case Index
        Case 1
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
                    MsgBox "Por favor introduzca el precio", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
                
                If Val(.Text) < 0 Then
                    MsgBox "El precio ingresado no es válido, por favor corrija la información", vbOKOnly, "Advertencia"
                    
                    .SetFocus
                    
                    Exit Sub
                End If
            End With
            
            vOrg_id = 490
            vOrder_type_id = 1133 'PRODUCCION
            'vOrder_type_id = 1113 'DESARROLLO
            vPrice_list_id = 363123
            
            vSold_from_org_id = 490
            
            vCreation_date = Date
            
            'vLine_type_id = 1112 'DESARROLLO
            vLine_type_id = 1132 'PRODUCCION
            
            vOrder_quantity_uom = "KG"
            vOrdered_quantity = Text2(0).Text
            
            Str = Combo2(2).Text
            ArrStr() = Split(Str, "[")
            
            vOrdered_item = Trim(Replace(ArrStr(1), "]", ""))
            vInventory_item_id = Replace(ArrStr(2), "]", "")
            vUnit_list_price = Text2(1).Text
             
            If vLine_number = 1 Then
                With RS
                    If .State = 1 Then .Close
                    .CursorLocation = adUseClient
                    .Open "Select count(*) + 1 from OE_ORDER_HEADERS_ALL", CN, adOpenStatic, adLockOptimistic
                
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        
                        vHeader_id = .Fields(0).Value
                        vOrder_Number = .Fields(0).Value
                        vCust_po_number = "VRN-" & .Fields(0).Value
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
                    .Fields(18) = vLote
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
                
                For I = 0 To 6
                    .Columns(I).Width = .Width / 5.5
                Next I
                
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
            
            For I = 0 To 1
                With Text2(I)
                    .Text = ""
                End With
            Next I
            
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
                
                For I = 0 To 6
                    .Columns(I).Width = .Width / 5.5
                Next I
                
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
            Else
                Text2(1).Text = ""
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
