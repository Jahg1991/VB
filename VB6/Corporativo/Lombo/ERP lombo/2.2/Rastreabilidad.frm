VERSION 5.00
Begin VB.Form frmRastreabilidad 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecutar Reporte"
   ClientHeight    =   2925
   ClientLeft      =   135
   ClientTop       =   480
   ClientWidth     =   4335
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3855
         Begin VB.ComboBox Combo1 
            Height          =   420
            Index           =   1
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   2415
         End
         Begin VB.ComboBox Combo1 
            Height          =   420
            Index           =   0
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command1 
            Height          =   615
            Left            =   1200
            Picture         =   "Rastreabilidad.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Folio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Origen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmRastreabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs          As New adodb.Recordset
    Dim Rs1         As New adodb.Recordset
    Dim Rsreporte1  As New adodb.Recordset
    
    '//OTROS
    Dim i           As Long

    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        Combo1(0).AddItem "Compra"
        Combo1(0).AddItem "Venta"
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Combo1_Click(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 0
                Combo1(1).Clear
                
                If Combo1(0).Text = "Compra" Then
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select folio from PO_HEADERS_ALL_P order by CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT) desc;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        
                        If .RecordCount > 0 Then
                            .MoveFirst
                            
                            While Not .EOF
                                Combo1(1).AddItem .Fields(0).Value
                                
                                .MoveNext
                            Wend
                        End If
                    End With
                Else
                    With Rs
                        If .State = 1 Then .Close
                        .CursorLocation = adodb.CursorLocationEnum.adUseClient
                        .Open "Select folio from PO_HEADERS_ALL_R order by CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT) desc;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                        .Requery
                        
                        If .RecordCount > 0 Then
                            .MoveFirst
                            
                            While Not .EOF
                                Combo1(1).AddItem .Fields(0).Value
                                
                                .MoveNext
                            Wend
                        End If
                    End With
                End If
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Combo1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command1_Click()
        On Error GoTo errHandler
        
        If Combo1(1).Text <> "" Then
            With Rsreporte1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "SELECT CASE WHEN t2.folio LIKE 'C-%' THEN '1. Compras' WHEN t2.folio LIKE 'P-%' THEN '2. Produccion' WHEN t2.folio LIKE 'V-%' THEN '3. Ventas' END tipo, " & _
                      "       t2.folio, " & _
                      "       t2.lote, " & _
                      "       convert(varchar,t2.fecha,105)AS fecha, " & _
                      "       dbo.initcap((SELECT t1.nombre FROM po_headers_all_v t1 WHERE t2.folio = t1.folio))AS nombre, " & _
                      "       upper(t2.itemcodigo)AS itemcodigo, " & _
                      "       dbo.initcap(t2.itemdescricion)AS itemdescricion, " & _
                      "       t2.tipotransaccion, " & _
                      "       t2.cantidad, " & _
                      "       t2.udm " & _
                      "FROM mtl_material_transactions t2 " & _
                      "WHERE t2.lote IN( " & _
                      "        SELECT DISTINCT lote FROM mtl_material_transactions WHERE cancelado = 'No' " & _
                      "            AND folio IN(SELECT DISTINCT folio FROM mtl_material_transactions WHERE cancelado = 'No' " & _
                      "                    AND lote IN(SELECT DISTINCT lote FROM mtl_material_transactions WHERE folio = '" & Combo1(1).Text & "'))) " & _
                      "  AND cancelado = 'No' " & _
                      "ORDER BY 1,5,CAST(replace(replace(replace(folio,'C-',''),'P-',''),'V-','')AS INT),3,7;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            End With
            
            If Rsreporte1.RecordCount <> 0 Then
                Unload Rastreabilidad
                
                With Rastreabilidad
                    Set .DataSource = Rsreporte1
                    
                    With .Sections("Section4")
                        .Controls("Label7").Caption = Combo1(0).Text & " Folio: " & Combo1(1).Text
                    End With
                    
                    With .Sections("Section1")
                        .Controls("Text1").DataField = "tipo"
                        .Controls("Text2").DataField = "nombre"
                        .Controls("Text3").DataField = "folio"
                        .Controls("Text4").DataField = "fecha"
                        .Controls("Text5").DataField = "itemdescricion"
                        .Controls("Text6").DataField = "tipotransaccion"
                        .Controls("Text7").DataField = "cantidad"
                        .Controls("Text8").DataField = "udm"
                        .Controls("Text9").DataField = "lote"
                    End With
                    
                    .Show 1
                End With
            End If
            
            Unload Me
        Else
            MsgBox "Seleccionar un folio", vbCritical, "Advertencia"
        End If
        Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Command1_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Salir_Click()
        On Error GoTo errHandler
        
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs1.Close
        If Rsreporte1.State = 1 Then Rsreporte1.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Rsreporte1 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmRastreabilidad:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
