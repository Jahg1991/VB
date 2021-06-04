VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuardar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   1920
      Picture         =   "frmGuardar.frx":324A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmGuardar.frx":3D4B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "¿Desea imprimir el informe?"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            DataReport1.Sections("Sección2").Controls("Etiqueta41").Caption = VarNombre
            DataReport1.Sections("Sección2").Controls("Etiqueta42").Caption = VarFecha_nacimiento
            DataReport1.Sections("Sección2").Controls("Etiqueta43").Caption = VarGenero
            DataReport1.Sections("Sección1").Controls("Etiqueta44").Caption = VarPeso
            DataReport1.Sections("Sección1").Controls("Etiqueta45").Caption = VarTalla
            DataReport1.Sections("Sección1").Controls("Etiqueta46").Caption = VarTension_arterial
            DataReport1.Sections("Sección1").Controls("Etiqueta47").Caption = VarVacuna_toxoide
            DataReport1.Sections("Sección1").Controls("Etiqueta48").Caption = VarOtras_vacunas
            DataReport1.Sections("Sección1").Controls("Etiqueta68").Caption = VarObservaciones_somatometria
            DataReport1.Sections("Sección1").Controls("Etiqueta58").Caption = VarColesterol
            DataReport1.Sections("Sección1").Controls("Etiqueta59").Caption = VarTrigliceridos
            DataReport1.Sections("Sección1").Controls("Etiqueta60").Caption = VarGlucosa
            DataReport1.Sections("Sección1").Controls("Etiqueta61").Caption = VarObservaciones_laboratorio
            DataReport1.Sections("Sección1").Controls("Etiqueta49").Caption = VarLavado_oidos
            DataReport1.Sections("Sección1").Controls("Etiqueta50").Caption = VarPrueba_audicion
            DataReport1.Sections("Sección1").Controls("Etiqueta51").Caption = VarObservaciones_audiometria
            DataReport1.Sections("Sección1").Controls("Etiqueta52").Caption = VarCardiologia
            DataReport1.Sections("Sección1").Controls("Etiqueta53").Caption = VarLimpieza_dental
            DataReport1.Sections("Sección1").Controls("Etiqueta54").Caption = VarRevision_dental
            DataReport1.Sections("Sección1").Controls("Etiqueta55").Caption = VarObservaciones_dental
            DataReport1.Sections("Sección1").Controls("Etiqueta57").Caption = VarDoccu
            DataReport1.Sections("Sección1").Controls("Etiqueta56").Caption = VarDocm
            DataReport1.Sections("Sección1").Controls("Etiqueta62").Caption = Varmastografia
            DataReport1.Sections("Sección1").Controls("Etiqueta63").Caption = VarConsulta_nutricion
            DataReport1.Sections("Sección1").Controls("Etiqueta64").Caption = VarPlatica_nutricion
            DataReport1.Sections("Sección1").Controls("Etiqueta65").Caption = VarObservaciones_nutricion
            DataReport1.Sections("Sección1").Controls("Etiqueta66").Caption = VarObservaciones_optometria
            DataReport1.Sections("Sección1").Controls("Etiqueta67").Caption = VarObservaciones_tuberculosis
            DataReport1.Show
            Form1.Enabled = True
            Unload Form4
        Case 1
            Form1.Enabled = True
            Unload Form4
    End Select
End Sub
