VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormImportarSomatometria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar a Somatometr�a"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   7845
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
   Icon            =   "Importar Somatometria.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   3840
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   7575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Importar"
      Height          =   555
      Left            =   3322
      TabIndex        =   2
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Personas que se van a importar:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Seleccionar 
         Caption         =   "Seleccionar"
         Shortcut        =   ^S
      End
      Begin VB.Menu Importar 
         Caption         =   "Importar"
         Shortcut        =   ^I
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FormImportarSomatometria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArchivoExcel As New Excel.Application
Dim intLoopCounter As Integer
Private Sub Command1_Click()
    On Error Resume Next
    List1.Clear
    CommonDialog1.DialogTitle = "Elige un archivo"
    CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then
        Exit Sub
    Else
        Label1.Caption = CommonDialog1.FileName
        ArchivoExcel.Workbooks.Open CommonDialog1.FileName
        ArchivoExcel.Worksheets(1).Activate
        With ArchivoExcel
            .Workbooks.Open CommonDialog1.FileName
            .Worksheets(1).Activate
            .Workbooks(1).Worksheets(1).Select
            For intLoopCounter = 1 To CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
                List1.AddItem .Range("B" & intLoopCounter)
            Next intLoopCounter
            .Workbooks.Close
            .Quit
        End With
        Set ArchivoExcel = Nothing
    End If
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    With ArchivoExcel
        .Workbooks.Open CommonDialog1.FileName
        .Worksheets(1).Activate
        .Workbooks(1).Worksheets(1).Select
        For intLoopCounter = 1 To CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
            With RsSomatometria
                .Requery
                .AddNew
                    .Fields("ID_AST") = ArchivoExcel.Cells(intLoopCounter, 1)
                    .Fields("NOMBRE") = ArchivoExcel.Cells(intLoopCounter, 2)
                    .Fields("FE_NAC") = ArchivoExcel.Cells(intLoopCounter, 3)
                    .Fields("EDAD") = ArchivoExcel.Cells(intLoopCounter, 4)
                    .Fields("GENERO") = ArchivoExcel.Cells(intLoopCounter, 5)
                    .Fields("TRAB_E") = ArchivoExcel.Cells(intLoopCounter, 6)
                    .Fields("AREA_T") = ArchivoExcel.Cells(intLoopCounter, 7)
                    .Fields("ID_EMP") = ArchivoExcel.Cells(intLoopCounter, 8)
                    .Fields("PARENT") = ArchivoExcel.Cells(intLoopCounter, 9)
                    .Fields("PES_KG") = ArchivoExcel.Cells(intLoopCounter, 10)
                    .Fields("TAL_MT") = ArchivoExcel.Cells(intLoopCounter, 11)
                    .Fields("TA") = ArchivoExcel.Cells(intLoopCounter, 12)
                    .Fields("VAC_TX") = ArchivoExcel.Cells(intLoopCounter, 13)
                    .Fields("VAC_OT") = ArchivoExcel.Cells(intLoopCounter, 14)
                    .Fields("OBSERV") = ArchivoExcel.Cells(intLoopCounter, 15)
                .Update
                .Requery
            End With
        Next intLoopCounter
        Set ArchivoExcel = Nothing
    End With
    MsgBox "Importaci�n completada", vbOKOnly, "Finalizado"
    Label1.Caption = ""
    List1.Clear
    CommonDialog1.FileName = ""
    Unload Me
    Form1.Enabled = True
End Sub
Private Sub Form_Load()
    On Error Resume Next
    With RsSomatometria
        If .State = 1 Then .Close
            .Open "Select * from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
End Sub
Private Sub Importar_Click()
    On Error Resume Next
    With ArchivoExcel
        .Workbooks.Open CommonDialog1.FileName
        .Worksheets(1).Activate
        .Workbooks(1).Worksheets(1).Select
        For intLoopCounter = 1 To CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
            With RsSomatometria
                .Requery
                .AddNew
                    .Fields("ID_AST") = ArchivoExcel.Cells(intLoopCounter, 1)
                    .Fields("NOMBRE") = ArchivoExcel.Cells(intLoopCounter, 2)
                    .Fields("FE_NAC") = ArchivoExcel.Cells(intLoopCounter, 3)
                    .Fields("EDAD") = ArchivoExcel.Cells(intLoopCounter, 4)
                    .Fields("GENERO") = ArchivoExcel.Cells(intLoopCounter, 5)
                    .Fields("TRAB_E") = ArchivoExcel.Cells(intLoopCounter, 6)
                    .Fields("AREA_T") = ArchivoExcel.Cells(intLoopCounter, 7)
                    .Fields("ID_EMP") = ArchivoExcel.Cells(intLoopCounter, 8)
                    .Fields("PARENT") = ArchivoExcel.Cells(intLoopCounter, 9)
                    .Fields("PES_KG") = ArchivoExcel.Cells(intLoopCounter, 10)
                    .Fields("TAL_MT") = ArchivoExcel.Cells(intLoopCounter, 11)
                    .Fields("TA") = ArchivoExcel.Cells(intLoopCounter, 12)
                    .Fields("VAC_TX") = ArchivoExcel.Cells(intLoopCounter, 13)
                    .Fields("VAC_OT") = ArchivoExcel.Cells(intLoopCounter, 14)
                    .Fields("OBSERV") = ArchivoExcel.Cells(intLoopCounter, 15)
                .Update
                .Requery
            End With
        Next intLoopCounter
        Set ArchivoExcel = Nothing
    End With
    MsgBox "Importaci�n completada", vbOKOnly, "Finalizado"
    Label1.Caption = ""
    CommonDialog1.FileName = ""
    List1.Clear
    Unload Me
    Form1.Enabled = True
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    Form1.Enabled = True
    Unload Me
End Sub
Private Sub Seleccionar_Click()
    On Error Resume Next
    List1.Clear
    CommonDialog1.DialogTitle = "Elige un archivo"
    CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then
        MsgBox "No se seleccion� ningun archivo", vbOKOnly, "Atenci�n"
        Exit Sub
    Else
        Label1.Caption = CommonDialog1.FileName
        With ArchivoExcel
            .Workbooks.Open CommonDialog1.FileName
            .Worksheets(1).Activate
            .Workbooks(1).Worksheets(1).Select
            For intLoopCounter = 1 To CInt(.ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
                List1.AddItem .Range("B" & intLoopCounter)
            Next intLoopCounter
            .Workbooks.Close
            .Quit
        End With
        Set ArchivoExcel = Nothing
    End If
End Sub
