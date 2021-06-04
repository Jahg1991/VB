VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1950
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1152.124
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   3000
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
      Connect         =   $"frmLogin.frx":324A
      OLEDBString     =   $"frmLogin.frx":32D2
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
   Begin VB.ComboBox Com_tipo 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txt_Usuario 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdEntrar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txt_Password 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Tipo:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Nombre de usuario:"
      Height          =   390
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
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
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  
' Flag
Dim OK As Boolean
  
  
Private Sub cmdEntrar_Click()
On Error Resume Next
      
    ' Cadena de conexión ( INDICAR EL PATH DE LA BASE DE DATOS )
    Const C_CADENA = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                     "Data Source=" & "C:\JAHG Software\Venta de cerdos\Databases\DB.mdb" & ";"
      
    ' Variable para el recordset
    Dim Rst_Login As Recordset
      
    ' crea el recordset
    Set Rst_Login = New Recordset
  
    Dim SQL As String
      
    ' consulta SQL ( Campos: Nombre y Password) _
                    Textbox ( txt_Usuario y txt_Password) _
                    Combobox (com_tipo) _
Tabla:                     Usuarios
      
    SQL = "SELECT Nombre, Tipo, Password " & _
                "FROM Usuarios " & _
                "WHERE Nombre = '" & txt_Usuario.Text & "'" _
                   & "AND Password = '" & txt_Password.Text & "'" _
                          & "AND Tipo = '" & Com_tipo.Text & "'"
  
    With Rst_Login
        ' Abre el recordset
        .Open SQL, C_CADENA
      
        ' Si el recordset está vacío es por que es incorrecto
        If .EOF Then
            MsgBox " El usuario o Password es incorrecto ", _
                     vbCritical, " Login incorrecto "
            ' Cierra y descarga el Recordset
            Rst_Login.Close
            Set Rst_Login = Nothing
            Exit Sub
        End If
    End With
    
        ' Control para permisos a dar de alta nuevos usuarios
   
               If Com_tipo.Text <> "Administrador" Then
                      FrmPrincipal.Nuevo_usuario.Enabled = False
                      Unload frmNuevo_usuario
               End If
    
      ' Control para permisos para captura
    If Com_tipo.Text = "Consulta" Then
              FrmPrincipal.Captura.Enabled = False
              Unload frmCaptura
              Unload frmNuevo_usuario
    End If
  
    ' Cierra y descarga el Recordset
    Rst_Login.Close
    Set Rst_Login = Nothing
      
    'Cambia el Flag para que no cierre el programa con End
    OK = True
      
    ' Descarga el formulario y prosigue en el SubMain
    Unload Me
End Sub
  
Private Sub cmdSalir_Click()
On Error Resume Next
    OK = False
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
     Com_tipo.AddItem "Administrador"
     Com_tipo.AddItem "Consulta"
     Com_tipo.AddItem "Captura"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set frmLogin = Nothing
    If OK = False Then
       End
    End If
End Sub
