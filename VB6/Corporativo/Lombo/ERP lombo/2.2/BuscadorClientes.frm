VERSION 5.00
Begin VB.Form frmBuscadorClientes 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00854E1B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404000&
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6660
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6375
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   2520
            Picture         =   "BuscadorClientes.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2400
            Width           =   1455
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   6135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1080
            TabIndex        =   2
            Top             =   120
            Width           =   5175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar"
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
            Index           =   2
            Left            =   -1080
            TabIndex        =   3
            Top             =   120
            Width           =   2055
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmBuscadorClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Dim Rs As New ADODB.Recordset
    
    Sub Form_Load()
        On Error GoTo errHandler
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .Open "Select * from HZ_PARTY where Cliente = 'Si' order by 1;", Cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic
            .Requery
            
            If .RecordCount <> 0 Then
                .MoveFirst
                
                List1.Clear
                
                While Not .EOF
                    List1.AddItem .Fields(1).Value
                    
                    .MoveNext
                Wend
            Else
                Unload Me
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorClientes:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Text1_Change()
        On Error GoTo errHandler
        
        List1.Clear
        
        With Rs
            If Text1 = "" Then
                .Filter = ""
                .Requery
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        List1.AddItem .Fields(1).Value
                        
                        .MoveNext
                    Wend
                End If
            Else
                .Filter = "nombre like '*" & Text1 & "*' or [Correo Electronico] = '" & Text1 & "'"
                .Requery
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        List1.AddItem .Fields(1).Value
                        
                        .MoveNext
                    Wend
                End If
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorClientes:Text1_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Command2_Click()
        On Error GoTo errHandler
        
        If List1.Text = "" Then
            MsgBox "Seleccione algún cliente", vbOKOnly, "Información"
        Else
            If TipoBusquedaCliente = "Venta" Then
                frmVentas.Combo1(0).Text = List1.Text
            Else
                frmPedidos.Combo1(0).Text = List1.Text
            End If
            
            Unload Me
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorClientes:Command2_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub List1_DblClick()
        On Error GoTo errHandler
        
        If List1.Text = "" Then
            MsgBox "Seleccione algún cliente", vbOKOnly, "Información"
        Else
            If TipoBusquedaCliente = "Venta" Then
                frmVentas.Combo1(0).Text = List1.Text
            Else
                frmPedidos.Combo1(0).Text = List1.Text
            End If
            
            Unload Me
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorClientes:List1_DblClick" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmAjusteInventario:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        
        Set frmBuscadorClientes = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmBuscadorClientes:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
