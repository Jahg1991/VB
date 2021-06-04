VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMod_elim_ventas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar o eliminar ventas"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17895
   ControlBox      =   0   'False
   Icon            =   "frmMod_elim_ventas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   17895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10200
      Top             =   8520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "VC"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Fecha_final 
      Height          =   285
      Left            =   8400
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Fecha_inicial 
      Height          =   285
      Left            =   7440
      TabIndex        =   18
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Buscar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Quitar filtro"
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   63111169
      CurrentDate     =   41508
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   8760
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   8880
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Primera venta"
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Venta anterior"
      Height          =   495
      Left            =   12120
      TabIndex        =   5
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Siguiente venta"
      Height          =   495
      Left            =   13080
      TabIndex        =   4
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Última venta"
      Height          =   495
      Left            =   14040
      TabIndex        =   3
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar venta"
      Height          =   495
      Left            =   16920
      TabIndex        =   2
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar cambios"
      Height          =   495
      Left            =   15960
      TabIndex        =   1
      Top             =   8280
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   9120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   63111169
      CurrentDate     =   41508
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMod_elim_ventas.frx":324A
      Height          =   8055
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   17640
      _ExtentX        =   31115
      _ExtentY        =   14208
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "Id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FECHA"
         Caption         =   "FECHA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "GRANJA"
         Caption         =   "GRANJA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "NO"
         Caption         =   "NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "KGS"
         Caption         =   "KGS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PROMEDIO"
         Caption         =   "PROMEDIO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "$KG"
         Caption         =   "$KG"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "SUBTOTAL"
         Caption         =   "SUBTOTAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "GUIAS"
         Caption         =   "GUIAS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "COMISIONES"
         Caption         =   "COMISIONES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "TOTAL"
         Caption         =   "TOTAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "CLIENTE"
         Caption         =   "CLIENTE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "TEJABAN"
         Caption         =   "TEJABAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "MORTALIDAD"
         Caption         =   "MORTALIDAD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "OBSERVACION"
         Caption         =   "OBSERVACION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2055,118
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1950,236
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar por:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   8760
      Width           =   855
   End
End
Attribute VB_Name = "frmMod_elim_ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub Combo1_Click()
On Error Resume Next

Select Case Combo1.Text

Case "Cliente"
Text1.Visible = True
DTPicker1.Visible = False
Combo2.Visible = False
Command7.Visible = True
Command8.Visible = False
Label2.Caption = "Cliente"
DTPicker2.Visible = False
Label3.Visible = False

Case "Granja"
Text1.Visible = False
DTPicker1.Visible = False
Combo2.Visible = True
Command7.Visible = True
Command8.Visible = False
Label2.Caption = "Granja"
DTPicker2.Visible = False
Label3.Visible = False

Case "Fecha"
Text1.Visible = False
DTPicker1.Visible = True
DTPicker2.Visible = True
Combo2.Visible = False
Command7.Visible = True
Command8.Visible = True
Label2.Caption = "Fecha"
Label3.Visible = True
End Select

End Sub

Private Sub Combo2_Click()
On Error Resume Next
With RS
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
Adodc1.Recordset.Filter = "GRANJA LIKE '*" & Combo2.Text & "*'"
Else
        ' Si el textbox no tiene nada, ... se limpia el Filtro
        Adodc1.Recordset.Filter = ""
        
        ' Vuelve a mostrar todos los registros en el dataGRid
        Set DataGrid1.DataSource = Adodc1.Recordset
        
        ' Opcional . Mueve el recordset al primer registro
        Adodc1.Recordset.MoveFirst
End If
End With
End Sub

Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.Update
End Sub

Private Sub Command2_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Recordset.Update
End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command5_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
End Sub


Private Sub Command7_Click()
On Error Resume Next
Text1.Text = ""
Combo2.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Adodc1.Recordset.Filter = ""
Set DataGrid1.DataSource = Adodc1.Recordset
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command8_Click()
On Error Resume Next
Fecha_inicial.Text = DTPicker1.Value
Fecha_final.Text = DTPicker2.Value

Command8.Enabled = False
End Sub

Private Sub Command9_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
 With Adodc1.Recordset
 If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
 .Sort = .Fields(ColIndex).[Name] & " Desc"
 Else
 .Sort = .Fields(ColIndex).[Name] & " Asc"
 End If
 End With
 End Sub

Private Sub DTPicker1_CloseUp()
On Error Resume Next
Command8.Enabled = True
End Sub

Private Sub DTPicker2_CloseUp()
On Error Resume Next
Command8.Enabled = True
End Sub

Private Sub Fecha_final_Change()

On Error Resume Next
With RS
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
Adodc1.Recordset.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
Else
        ' Si el textbox no tiene nada, ... se limpia el Filtro
        Adodc1.Recordset.Filter = ""
        
        ' Vuelve a mostrar todos los registros en el dataGRid
        Set DataGrid1.DataSource = Adodc1.Recordset
        
        ' Opcional . Mueve el recordset al primer registro
        Adodc1.Recordset.MoveFirst
End If
End With
End Sub

Private Sub Fecha_inicial_Change()
On Error Resume Next
With RS
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
Adodc1.Recordset.Filter = "FECHA >= " & _
"# " + Fecha_inicial.Text + " # And  FECHA <= # " + Fecha_final.Text + " #"
Else
        ' Si el textbox no tiene nada, ... se limpia el Filtro
        Adodc1.Recordset.Filter = ""
        
        ' Vuelve a mostrar todos los registros en el dataGRid
        Set DataGrid1.DataSource = Adodc1.Recordset
        
        ' Opcional . Mueve el recordset al primer registro
        Adodc1.Recordset.MoveFirst
End If
End With
End Sub

Private Sub Form_Load()
On Error Resume Next
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;"
Set RS = cn.Execute("SELECT * FROM VC")

Combo1.AddItem "Cliente"
Combo1.AddItem "Granja"
Combo1.AddItem "Fecha"

Combo2.AddItem "Terrero"
Combo2.AddItem "Isabel"
Combo2.AddItem "Laja"
Combo2.AddItem "Cuna"
Combo2.AddItem "Sapo"
Combo2.AddItem "Moro"
Combo2.AddItem "Loma"

DTPicker1.Value = Date
DTPicker2.Value = Date

End Sub

Private Sub Text1_Change()
On Error Resume Next
With RS
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
Adodc1.Recordset.Filter = "CLIENTE LIKE '*" & Text1 & "*'"
Else
        ' Si el textbox no tiene nada, ... se limpia el Filtro
        Adodc1.Recordset.Filter = ""
        
        ' Vuelve a mostrar todos los registros en el dataGRid
        Set DataGrid1.DataSource = Adodc1.Recordset
        
        ' Opcional . Mueve el recordset al primer registro
        Adodc1.Recordset.MoveFirst
End If
End With
End Sub
