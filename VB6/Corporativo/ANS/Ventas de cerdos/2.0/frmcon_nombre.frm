VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmcon_nombre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "B?squeda por cliente"
   ClientHeight    =   11475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17895
   ControlBox      =   0   'False
   Icon            =   "frmcon_nombre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11475
   ScaleWidth      =   17895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picImprimir 
      Height          =   255
      Left            =   6360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   5640
      Picture         =   "frmcon_nombre.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Imprimir 
      Height          =   375
      Left            =   5040
      Picture         =   "frmcon_nombre.frx":387D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   16560
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quitar filtro"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"frmcon_nombre.frx":3E2C
      OLEDBString     =   $"frmcon_nombre.frx":3EB4
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcon_nombre.frx":3F3C
      Height          =   10575
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   17640
      _ExtentX        =   31115
      _ExtentY        =   18653
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
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmcon_nombre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next
Text1.Text = ""
Adodc1.Recordset.Filter = ""
Set DataGrid1.DataSource = Adodc1.Recordset
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
frmNombre_archivo.Show
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

Private Sub Form_Load()
On Error Resume Next
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;"
Set RS = cn.Execute("SELECT * FROM VC")

DataGrid1.AllowAddNew = False 'para no agregar registros nuevos
DataGrid1.AllowUpdate = False 'para no modificar los registros existentes
End Sub

Private Sub Imprimir_Click()
On Error Resume Next

Printer.Orientation = vbPRORLandscape
Printer.PaperSize = vbPRPSLetter 'Tipo de Papel

picImprimir.Picture = CaptureClient(Me)
Printer.PaintPicture picImprimir.Picture, 0, 0, Printer.ScaleWidth, (Me.ScaleHeight * Printer.ScaleWidth) / Me.ScaleWidth, , , Me.ScaleWidth, Me.ScaleHeight
Printer.EndDoc

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
