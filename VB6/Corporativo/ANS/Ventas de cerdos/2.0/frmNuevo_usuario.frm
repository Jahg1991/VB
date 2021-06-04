VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmNuevo_usuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo usuario"
   ClientHeight    =   5205
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "frmNuevo_usuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075.286
   ScaleMode       =   0  'User
   ScaleWidth      =   5732.265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   4080
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Usuarios"
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
      Bindings        =   "frmNuevo_usuario.frx":324A
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
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
         DataField       =   "Password"
         Caption         =   "Password"
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
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1633,677
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   390
      Left            =   2482
      TabIndex        =   11
      Top             =   4680
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quitar filtro"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   3480
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   3600
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   840
      Picture         =   "frmNuevo_usuario.frx":325F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      Picture         =   "frmNuevo_usuario.frx":378C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   4200
      TabIndex        =   5
      Top             =   720
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nombre de usuario:"
      Height          =   390
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Tipo:"
      Height          =   270
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Editor o eliminar usuarios"
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nombre de usuario:"
      Height          =   390
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmNuevo_usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub cmdCancel_Click()
On Error Resume Next
txtUserName = ""
txtPassword = ""
Combo1.Text = ""
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("Nombre") = txtUserName.Text
Adodc1.Recordset.Fields("Password") = txtPassword.Text
Adodc1.Recordset.Fields("Tipo") = Combo1.Text
Adodc1.Recordset.Update
txtUserName = ""
txtPassword = ""
Combo1.Text = ""
Unload Me
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
Text1.Text = ""
Adodc1.Recordset.Filter = ""
Set DataGrid1.DataSource = Adodc1.Recordset
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command4_Click()
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

Private Sub Form_Load()
On Error Resume Next

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=C:\JAHG Software\Venta de cerdos\Databases\DB.mdb;"
Set RS = cn.Execute("SELECT * FROM Usuarios")

Combo1.AddItem "Administrador"
Combo1.AddItem "Captura"
Combo1.AddItem "Consulta"
DataGrid1.AllowAddNew = False 'para no agregar registros nuevos
End Sub

Private Sub Text1_Change()
On Error Resume Next
With RS
On Error Resume Next ' porque me da error si en el textbox no hay nada
If Option1.Value = True Then
Adodc1.Recordset.Filter = "Nombre LIKE '*" & Text1 & "*'"
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
