VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10425
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   9840
      TabIndex        =   19
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   8760
      TabIndex        =   18
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox txt_Mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBC9D0&
      BorderStyle     =   0  'None
      ForeColor       =   &H0035434A&
      Height          =   5445
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   1200
      Width           =   8775
   End
   Begin VB.TextBox txt_Adjunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBC9D0&
      BorderStyle     =   0  'None
      ForeColor       =   &H0035434A&
      Height          =   405
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   720
      Width           =   8295
   End
   Begin VB.TextBox txt_Asunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00BBC9D0&
      BorderStyle     =   0  'None
      ForeColor       =   &H0035434A&
      Height          =   405
      Left            =   1440
      TabIndex        =   15
      Top             =   240
      Width           =   8775
   End
   Begin VB.TextBox txt_Para 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.TextBox txt_De 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1440
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.TextBox txt_Password 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   1680
      Width           =   8775
   End
   Begin VB.TextBox txt_Usuario 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1200
      Width           =   8775
   End
   Begin VB.TextBox txt_Puerto 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   720
      Width           =   8775
   End
   Begin VB.TextBox txt_Servidor 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   240
      Width           =   8775
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerpo"
      ForeColor       =   &H0035434A&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjunto"
      ForeColor       =   &H0035434A&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asunto"
      ForeColor       =   &H0035434A&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enviar a"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Envìa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ' El ejemplo para poder enviar el mail necesita la referencia a: _
      > Miscrosoft CDO Windows For 2000 Library ( es el archivo dll cdosys.dll )
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Private Function Enviar_Mail_CDO(SerVidor_SMTP As String, _
                                 Para As String, _
                                 De As String, _
                                 Asunto As String, _
                                 Mensaje As String, _
                                 Optional Path_Adjunto As String, _
                                 Optional Puerto As String = "465", _
                                 Optional Usuario As String, _
                                 Optional Password As String, _
                                 Optional Usar_Autentificacion As Boolean = True, _
                                 Optional Usar_SSL As Boolean = True) As Boolean
          
          
        Me.MousePointer = vbHourglass
          
        ' Variable de objeto Cdo.Message
        Dim Obj_Email As CDO.Message
                
          
        ' Crea un Nuevo objeto CDO.Message
        Set Obj_Email = New CDO.Message
          
        ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
          del servidor o su dirección IP )
        Obj_Email.Configuration.Fields(cdoSMTPServer) = SerVidor_SMTP
          
        Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
          
        ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
          465 o  el puerto 587 ( este último me dio error )
        Obj_Email.Configuration.Fields.Item _
            ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto)
      
          
        ' Indica el tipo de autentificación con el servidor de correo _
         El valor 0 no requiere autentificarse, el valor 1 es con autentificación
        Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                    "configuration/smtpauthenticate") = Abs(Usar_Autentificacion)
          
        ' Tiempo máximo de espera en segundos para la conexión
        Obj_Email.Configuration.Fields.Item _
            ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
      
          
        ' Configura las opciones para el login en el SMTP
        If Usar_Autentificacion Then
      
            ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
            mas el @gmail.com )
            Obj_Email.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Usuario
      
            ' Password de la cuenta
            Obj_Email.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password
      
            ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
            Obj_Email.Configuration.Fields.Item _
               ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL
          
        End If
          
      
        ' *********************************************************************************
        ' Estructura del mail
        '**********************************************************************************
          
        ' Dirección del Destinatario
        Obj_Email.To = Para
          
        ' Dirección del remitente
        Obj_Email.From = De
          
        ' Asunto del mensaje
        Obj_Email.Subject = Asunto
          
        ' Cuerpo del mensaje
        Obj_Email.TextBody = Mensaje
          
        'Ruta del archivo adjunto
          
        If Path_Adjunto <> vbNullString Then
            Obj_Email.AddAttachment (Path_Adjunto)
        End If
          
        ' Actualiza los datos antes de enviar
        Obj_Email.Configuration.Fields.Update
          
        On Error Resume Next
        ' Envía el email
        Obj_Email.Send
          
          
        If Err.Number = 0 Then
           Enviar_Mail_CDO = True
        Else
           MsgBox Err.Description, vbCritical, " Error al enviar el amil "
        End If
          
        ' Descarga la referencia
        If Not Obj_Email Is Nothing Then
            Set Obj_Email = Nothing
        End If
          
        On Error GoTo 0
        Me.MousePointer = vbNormal
      
    End Function
      
    Private Sub Command1_Click()
          
        Dim ret As Boolean
          
        ' Asegurarse de pasar bien los últimos dos parámetros _
         ( Si usa login y si el server usa SSL)
          
        ret = Enviar_Mail_CDO(txt_Servidor, _
                              txt_Para, _
                              txt_De, _
                              "Correo de " + VARUSUARIO + ", Asunto " + txt_Asunto, _
                              txt_Mensaje, _
                              txt_Adjunto, _
                              txt_Puerto, _
                              txt_Usuario, _
                              txt_Password, _
                              True, _
                              True)
          
        ' Si devuelve true es por que no hubo errores en el envio
        If ret Then
            MsgBox " .. Maneje enviado ", vbInformation
            Unload Me
        End If
        
    End Sub
      
Private Sub Command2_Click()

    CommonDialog1.ShowOpen
    txt_Adjunto = CommonDialog1.FileName
    
End Sub

    Private Sub Form_Load()
      
        Me.Caption = "Enviar correo externo a " + Form3.Label2(4).Caption
        Command1.Caption = " Enviar mail "
          
        txt_Servidor.Text = "smtp.gmail.com"
        txt_Servidor.Visible = False
        txt_Para = Form3.Label2(4).Caption
        txt_De = "anssoporte07@gmail.com"
        txt_Asunto = ""
        txt_Mensaje = ""
        txt_Adjunto = vbNullString
        txt_Puerto.Text = 465
        txt_Puerto.Visible = False
        txt_Password = "Santateresa1"
        txt_Password.Visible = False
        txt_Usuario = "anssoporte07@gmail.com"
        txt_Usuario.Visible = False
        
    End Sub

