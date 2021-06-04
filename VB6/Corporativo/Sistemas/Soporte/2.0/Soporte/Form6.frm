VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar agenda"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar archivo..."
      Height          =   615
      Index           =   2
      Left            =   6000
      Picture         =   "Form6.frx":72FA
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "Form6.frx":7C53
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArchivoExcel As New Excel.Application
Dim intLoopCounter As Integer
Dim CnDb As New ADODB.Connection
Dim RsSales As New Recordset

Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 2
            CommonDialog1.DialogTitle = "Elige un archivo"
            CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName = "" Then
                Exit Sub
            Else
                Text1.Text = CommonDialog1.FileName
            End If

        Case 3
            With ArchivoExcel
                .Workbooks.Open CommonDialog1.FileName
                .Worksheets(1).Activate
                .Workbooks(1).Worksheets(1).Select
                For intLoopCounter = 2 To CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
                    With RsSales
                        .Requery
                        .AddNew
                        .Fields("NOMBRE") = ArchivoExcel.Cells(intLoopCounter, 1)
                        .Fields("EXTENSION") = ArchivoExcel.Cells(intLoopCounter, 2)
                        .Fields("CELULAR") = ArchivoExcel.Cells(intLoopCounter, 3)
                        .Fields("RADIO") = ArchivoExcel.Cells(intLoopCounter, 4)
                        .Fields("INTERNO") = ArchivoExcel.Cells(intLoopCounter, 5)
                        .Fields("EXTERNO") = ArchivoExcel.Cells(intLoopCounter, 6)
                        .Update
                        .Requery
                    End With
                Next intLoopCounter
                Set ArchivoExcel = Nothing
            End With
            MsgBox "Importación completada", vbOKOnly, "Finalizado"
            
            Text1 = ""
            Command1(2).SetFocus
            
    End Select
    
End Sub

Private Sub Form_Load()

    On Error Resume Next

    With RsSales
        If .State = 1 Then .Close
           .Open "Select * from AGENDA", CnDb, adOpenStatic, adLockOptimistic
           .Requery
    End With
    
    With CnDb
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=Soporte1;Persist Security Info=True;User ID=Soporte;Initial Catalog=SOPORTE;Data Source=SQLSERVER\SQLEXPRESS;"
    End With
    
End Sub
