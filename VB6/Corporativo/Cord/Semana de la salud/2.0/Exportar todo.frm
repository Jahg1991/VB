VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormExportarTodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar información"
   ClientHeight    =   1500
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
   Icon            =   "Exportar todo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar"
      Height          =   555
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar como..."
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu GuardarComo 
         Caption         =   "Guardar como..."
         Shortcut        =   ^G
      End
      Begin VB.Menu Exportar 
         Caption         =   "Exportar"
         Shortcut        =   ^E
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FormExportarTodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        MsgBox "Selecciona donde quieres guardar el archivo", vbOKOnly, "Atención"
        Exit Sub
    Else
        Text1.Text = CommonDialog1.FileName
    End If
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim exportFileName As String
    Dim i As Long
    Dim j As Long
    Dim numFilas As Long
    i = 1
    j = 1
    numFilas = 0
    exportFileName = FormExportarTodo.CommonDialog1.FileName
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    xlApp.DisplayAlerts = False
    While (numFilas < RsAllDate.Fields.Count)
        xlSheet.Cells(i, j) = RsAllDate.Fields(numFilas).Name
        j = j + 1
        numFilas = numFilas + 1
    Wend
        i = i + 1
    While (Not RsAllDate.EOF)
        j = 0
        While (j < numFilas)
            xlSheet.Cells(i, j + 1).Value = RsAllDate(j)
            j = j + 1
        Wend
        i = i + 1
        RsAllDate.MoveNext
    Wend
    xlSheet.SaveAs exportFileName
    xlBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    MsgBox "Archivo creado correctamente"
    Text1.Text = ""
    Unload Me
    Form1.Enabled = True
    Exit Sub
err:
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        MsgBox "El archivo no ha podido ser creado" & vbNewLine & err.Description
    Text1.Text = ""
    Unload Me
    Form1.Enabled = True
End Sub
Private Sub Exportar_Click()
    On Error Resume Next
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim exportFileName As String
    Dim i As Long
    Dim j As Long
    Dim numFilas As Long
    i = 1
    j = 1
    numFilas = 0
    exportFileName = FormExportarTodo.CommonDialog1.FileName
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    xlApp.DisplayAlerts = False
    While (numFilas < RsAllDate.Fields.Count)
        xlSheet.Cells(i, j) = RsAllDate.Fields(numFilas).Name
        j = j + 1
        numFilas = numFilas + 1
    Wend
        i = i + 1
    While (Not RsAllDate.EOF)
        j = 0
        While (j < numFilas)
            xlSheet.Cells(i, j + 1).Value = RsAllDate(j)
            j = j + 1
        Wend
        i = i + 1
        RsAllDate.MoveNext
    Wend
    xlSheet.SaveAs exportFileName
    xlBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    MsgBox "Archivo creado correctamente"
    Text1.Text = ""
    Unload Me
    Form1.Enabled = True
    Exit Sub
err:
        Set xlApp = Nothing
        Set xlBook = Nothing
        Set xlSheet = Nothing
        MsgBox "El archivo no ha podido ser creado" & vbNewLine & err.Description
    Text1.Text = ""
    Unload Me
    Form1.Enabled = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    With RsAllDate
        If .State = 1 Then .Close
            .Open "Select * from All_date", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
End Sub

Private Sub GuardarComo_Click()
    On Error Resume Next
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Archivo de excel 97-03|*.xls"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        MsgBox "Selecciona donde quieres guardar el archivo", vbOKOnly, "Atención"
        Exit Sub
    Else
        Text1.Text = ""
    End If
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    Unload Me
    Form1.Enabled = True
End Sub
