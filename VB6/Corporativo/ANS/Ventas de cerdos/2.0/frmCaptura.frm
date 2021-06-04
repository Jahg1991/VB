VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmCaptura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmCaptura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1853
      TabIndex        =   45
      Top             =   5760
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6375
      Left            =   120
      TabIndex        =   30
      Top             =   6840
      Width           =   4215
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1200
         Top             =   5400
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
      Begin VB.TextBox Text26 
         DataField       =   "FECHA"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   44
         Text            =   "0"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Text25 
         DataField       =   "GRANJA"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   43
         Text            =   "0"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text24 
         DataField       =   "NO"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   42
         Text            =   "0"
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text23 
         DataField       =   "KGS"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Text            =   "0"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text22 
         DataField       =   "PROMEDIO"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Text            =   "0"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Text21 
         DataField       =   "$KG"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Text            =   "0"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text20 
         DataField       =   "SUBTOTAL"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Text            =   "0"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox Text19 
         DataField       =   "GUIAS"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Text            =   "0"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text18 
         DataField       =   "COMISIONES"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Text            =   "0"
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox Text17 
         DataField       =   "TOTAL"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Text            =   "0"
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox Text16 
         DataField       =   "CLIENTE"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   34
         Text            =   "0"
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Text15 
         DataField       =   "TEJABAN"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Text            =   "0"
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox Text14 
         DataField       =   "MORTALIDAD"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Text            =   "0"
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox Text13 
         DataField       =   "OBSERVACION"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Text            =   "0"
         Top             =   4920
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1320
      TabIndex        =   28
      Top             =   5160
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1320
      TabIndex        =   27
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Format          =   63045633
      CurrentDate     =   41507
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   26
      Text            =   "-"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Text            =   "Ninguna"
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1320
      TabIndex        =   24
      Text            =   "0"
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1320
      TabIndex        =   23
      Text            =   "-"
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1320
      TabIndex        =   22
      Text            =   "-"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   20
      Text            =   "0"
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Text            =   "0"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Text            =   "0"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   16
      Text            =   "0"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Text            =   "0"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "0"
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Granja"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Número"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Kilogramos"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Peso Promedio"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Precio del kg"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Subtotal"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Guias"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Comisiones"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Tejaban"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Mortalidad"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciòn"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("FECHA") = DTPicker1.Value
Adodc1.Recordset.Fields("GRANJA") = Combo1.Text
Adodc1.Recordset.Fields("NO") = Text1.Text
Adodc1.Recordset.Fields("KGS") = Text2.Text
Adodc1.Recordset.Fields("PROMEDIO") = Text3.Text
Adodc1.Recordset.Fields("$KG") = Text4.Text
Adodc1.Recordset.Fields("SUBTOTAL") = Text5.Text
Adodc1.Recordset.Fields("GUIAS") = Text6.Text
Adodc1.Recordset.Fields("COMISIONES") = Text7.Text
Adodc1.Recordset.Fields("TOTAL") = Text8.Text
Adodc1.Recordset.Fields("CLIENTE") = Text9.Text
Adodc1.Recordset.Fields("TEJABAN") = Text10.Text
Adodc1.Recordset.Fields("MORTALIDAD") = Text11.Text
Adodc1.Recordset.Fields("OBSERVACION") = Text12.Text
Adodc1.Recordset.Update
MsgBox ("Venta guardada con exito")
Text1.Text = "0"
Text2.Text = "0"
Text3.Text = "0"
Text4.Text = "0"
Text5.Text = "0"
Text6.Text = "0"
Text7.Text = "0"
Text8.Text = "0"
Text9.Text = "-"
Text10.Text = "-"
Text11.Text = "0"
Text12.Text = "Ninguna"
Combo1.Text = "-"
DTPicker1.Value = Date
DTPicker1.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker1.Value = Date
Combo1.AddItem "Terrero"
Combo1.AddItem "Isabel"
Combo1.AddItem "Laja"
Combo1.AddItem "Cuna"
Combo1.AddItem "Sapo"
Combo1.AddItem "Moro"
Combo1.AddItem "Loma"
End Sub

Private Sub Text1_Change()
On Error Resume Next
Text3 = Val(Text2.Text) / Val(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text11.SetFocus
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text12.SetFocus
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
Text3 = Val(Text2.Text) / Val(Text1.Text)
Text5 = Val(Text2.Text) * Val(Text4.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
On Error Resume Next
Text5 = Val(Text2.Text) * Val(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text5_Change()
On Error Resume Next
Text8 = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
End Sub

Private Sub Text6_Change()
On Error Resume Next
Text8 = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text7_Change()
On Error Resume Next
Text8 = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text9.SetFocus
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Text10.SetFocus
End If
End Sub
