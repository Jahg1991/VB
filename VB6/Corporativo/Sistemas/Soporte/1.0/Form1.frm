VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Solicitud"
   ClientHeight    =   5580
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Ver agenda"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mis solicitudes"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nueva Solicitud"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6495
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   2468
         TabIndex        =   7
         Top             =   720
         Width           =   2055
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Oracle EBS"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Otro"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Impresoras"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Correo"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Archivos"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   5
         Top             =   3240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   3360
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6480
      Top             =   8160
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   7680
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
      Connect         =   $"Form1.frx":164A
      OLEDBString     =   $"Form1.frx":16D1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "solicitudes"
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
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "SOLICITANTE"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "ESTATUS"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "HORA"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "FECHA"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "IMAGEN"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "DETALLES"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "TIPO"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
  
' Declaración del Api GetUserName
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long) As Long
  
' Retorna un String con el nombre de usuario actual de windows
' ************************************************************
Private Function get_Usuario() As String
      
    Dim Nombre As String, ret As Long
      
    ' Buffer
    Nombre = Space$(250)
      
    ' Tamaño
    ret = Len(Nombre)
      
    If GetUserName(Nombre, ret) = 0 Then
        get_Usuario = vbNullString
    Else
        ' Extrae solo los caracteres
        get_Usuario = Left$(Nombre, ret - 1)
    End If
      
End Function

Private Sub Command1_Click()

    On Error Resume Next
    Adodc1.Recordset.Requery
    Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("TIPO") = Label3.Caption
        Adodc1.Recordset.Fields("DETALLES") = Text1.Text
        Adodc1.Recordset.Fields("FECHA") = Label2(1).Caption
        Adodc1.Recordset.Fields("HORA") = Label2(2).Caption
        Adodc1.Recordset.Fields("ESTATUS") = "Pendiente"
        Adodc1.Recordset.Fields("SOLICITANTE") = Label2(0).Caption
    Adodc1.Recordset.Update
    Adodc1.Recordset.Requery
    Unload Me
    Form1.Show
    
End Sub


Private Sub Command3_Click()

    Form3.Show

End Sub

Private Sub Command4_Click()

    Form4.Show
    
End Sub

Private Sub Form_Load()

    ' Muestra el usuario
    Label2(0) = get_Usuario
    
    ' Mestra la fecha
    Label2(1).Caption = Format(Now, "yyyy/mm/dd")
    
    ' Muetra la hora
    Label2(2) = Time
    
    ' Ponemos un tipo
    Option1(0).Value = True
    Label3.Caption = "Archivos"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Form3
    Unload Form4

End Sub

Private Sub Option1_Click(Index As Integer)
    
    ' Pasamos el tipo a text en un label
    Select Case Index
        Case 0
            Label3.Caption = "Archivos"
        Case 1
            Label3.Caption = "Correo"
        Case 2
            Label3.Caption = "Impresoras"
        Case 3
            Label3.Caption = "Oracle EBS"
        Case 4
            Label3.Caption = "Otro"
    End Select
    
End Sub

Private Sub Timer1_Timer()

    ' Muetra la hora
    Label2(2) = Time

End Sub
