VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PigSale v.1.0 - Importar ventas"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Sans Unicode"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Height          =   615
      Index           =   2
      Left            =   6000
      Picture         =   "PigSale v.1.0. - Importar ventas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
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
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   1
         Left            =   6840
         Picture         =   "PigSale v.1.0. - Importar ventas.frx":0959
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   0
         Left            =   4560
         Picture         =   "PigSale v.1.0. - Importar ventas.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   5
         Left            =   3000
         Picture         =   "PigSale v.1.0. - Importar ventas.frx":1665
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   4
         Left            =   1440
         Picture         =   "PigSale v.1.0. - Importar ventas.frx":201B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "PigSale v.1.0. - Importar ventas.frx":28E2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
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

Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    Select Case Index
    
        Case 0
            Unload Me
        
        Case 1
            Form1.Show
            Form1.Text1(0).Text = ""
            Form1.Text1(1).Text = ""
            Unload Me
    
        Case 2
            MsgBox "Para que la importación sea realizada adecuadamente se debe utilizar el formato ubicado en 'C:\JAHG Software\PigSale\Formato de importación.xls'", vbOKOnly, "Atención"
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
                        .Fields("FECHA") = ArchivoExcel.Cells(intLoopCounter, 1)
                        .Fields("GRANJA") = ArchivoExcel.Cells(intLoopCounter, 2)
                        .Fields("NUMERO") = ArchivoExcel.Cells(intLoopCounter, 3)
                        .Fields("KILOS") = ArchivoExcel.Cells(intLoopCounter, 4)
                        .Fields("PROMEDIO") = ArchivoExcel.Cells(intLoopCounter, 5)
                        .Fields("CLIENTE") = ArchivoExcel.Cells(intLoopCounter, 6)
                        .Fields("TEJABAN") = ArchivoExcel.Cells(intLoopCounter, 7)
                        .Fields("MORTANDAD") = ArchivoExcel.Cells(intLoopCounter, 8)
                        .Fields("OBSERVACIONES") = ArchivoExcel.Cells(intLoopCounter, 9)
                        .Fields("ANO") = ArchivoExcel.Cells(intLoopCounter, 10)
                        .Fields("SEMANA") = ArchivoExcel.Cells(intLoopCounter, 11)
                        .Update
                        .Requery
                    End With
                Next intLoopCounter
                Set ArchivoExcel = Nothing
            End With
            MsgBox "Importación completada", vbOKOnly, "Finalizado"
            
            Text1 = ""
            Command1(2).SetFocus

        Case 4
            Form4.Show
            Form4.Label1.Caption = Label1.Caption
            Unload Me
        
        Case 5
            Form2.Show
            Form2.Label1.Caption = Label1.Caption
            Unload Me
            
    End Select
    
End Sub

Private Sub Form_Load()

    On Error Resume Next

    With RsSales
        If .State = 1 Then .Close
           .Open "Select * from VENTAS", CnDb, adOpenStatic, adLockOptimistic
           .Requery
    End With
    
End Sub
