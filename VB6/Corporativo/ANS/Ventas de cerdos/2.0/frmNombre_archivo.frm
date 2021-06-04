VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmNombre_archivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombre del archivo"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6690
   ControlBox      =   0   'False
   Icon            =   "frmNombre_archivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   1200
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
      CommandType     =   8
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
      Connect         =   $"frmNombre_archivo.frx":324A
      OLEDBString     =   $"frmNombre_archivo.frx":32D2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmNombre_archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Exportar_ADO_Excel(sPathDB As String, SQL As String, sOutputPathXLS As String) As Boolean
On Error Resume Next
    
    On Error GoTo ErrSub
    
    Dim cn1         As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    
    Me.Enabled = False
    
   ' -- Abrir la base
    cn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"
        
    ' -- Abrir el Recordset pasándole la cadena sql
    rec.Open SQL, cn1
    
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
    
    Excel.Visible = True: Excel.UserControl = True
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
    Next
    
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rec
    Else

        arrData = rec.GetRows

        iRec = UBound(arrData, 2) + 1
        
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1

                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))

                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
            
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
    End If

    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit

    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
    cn1.Close
    
    Set rec = Nothing
    Set cn1 = Nothing
    ' -- guardar el libro
    Libro.saveAs sOutputPathXLS
    Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    Excel.quit
    Set Excel = Nothing
    
    Exportar_ADO_Excel = True
    Me.Enabled = True
    Exit Function
ErrSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel = False
    Me.Enabled = True
End Function
 ' Excel

Private Function GetData(vValue As Variant) As Variant
On Error Resume Next
    Dim x As Long, y As Long, xMax As Long, yMax As Long, T As Variant
    
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
    
    ReDim T(xMax, yMax)
    For x = 0 To xMax
        For y = 0 To yMax
            T(x, y) = vValue(y, x)
        Next y
    Next x
    
    GetData = T
End Function

Private Sub Command1_Click()
On Error Resume Next
Dim sPathDB        As String
    Dim Consulta    As String

    ' -- Path de la base de datos
    sPathDB = "C:\JAHG Software\Venta de cerdos\Databases\DB.MDB"

    ' -- Cadena Sql
    Consulta = "Select * From VC"

    ' -- Enviar el Path de la base de datos y la consulta sql
    If Exportar_ADO_Excel(sPathDB, Consulta, "C:\JAHG Software\Venta de cerdos\Reportes\" + Text1.Text + ".xls") Then
       MsgBox "Terminado", vbInformation
    End If
    Command1.Enabled = False
    Text1.Text = ""
    frmNombre_archivo.Hide
    
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Text1 = "" Then
Command1.Enabled = False
Else
If Text1 <> "" Then
Command1.Enabled = True
End If
End If
End Sub
