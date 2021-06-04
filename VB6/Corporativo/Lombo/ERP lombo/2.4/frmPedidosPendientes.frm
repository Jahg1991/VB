VERSION 5.00
Begin VB.Form frmPedidosPendientes 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Pendientes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17415
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6840
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   6
      Left            =   7080
      TabIndex        =   9
      Top             =   2760
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00E0E0E0&
         Height          =   2295
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   6495
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6495
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   2
      Left            =   7080
      TabIndex        =   3
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6495
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmPedidosPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'Nombre:        frmPedidosPendientes
'Proposito:     Visualizar turnos de venta
'
'Revisiones
'Version    Fecha          Nombre               Revision
'-----------------------------------------------------------------------------------
'1.0        13/05/2021     Alfredo Hernandez    Creacion
'
'***********************************************************************************
    Option Explicit
    
    '===============================================================================
    'DECLARACION DE VARIABLES
    '===============================================================================
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        With frmPedidosPendientes
            .WindowState = 2
        End With
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
        
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "SELECT top 4 min(id) as id,Folio, Nombre From PO_LINES_ALL Where Tipo= 'Pedidos' AND cancelado= 'No' group by Folio, Nombre order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount <> 0 Then
                .MoveFirst
                Label1(0).Caption = "TURNO #1" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                .MoveNext
                If Not .EOF Then
                    Label1(3).Caption = "TURNO #2" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                    .MoveNext
                    If Not .EOF Then
                        Label1(5).Caption = "TURNO #3" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                        .MoveNext
                        If Not .EOF Then
                            Label1(7).Caption = "TURNO #4" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                        Else
                            With Label1(7)
                                .Caption = "TURNO #4"
                            End With
                        End If
                    Else
                        With Label1(5)
                            .Caption = "TURNO #3"
                        End With
                        
                        With Label1(7)
                            .Caption = "TURNO #4"
                        End With
                    End If
                Else
                    With Label1(3)
                        .Caption = "TURNO #2"
                    End With
                    
                    With Label1(5)
                        .Caption = "TURNO #3"
                    End With
                    
                    With Label1(7)
                        .Caption = "TURNO #4"
                    End With
                End If
            Else
                With Label1(0)
                    .Caption = "TURNO #1"
                End With
                
                With Label1(3)
                    .Caption = "TURNO #2"
                End With
                
                With Label1(5)
                    .Caption = "TURNO #3"
                End With
                
                With Label1(7)
                    .Caption = "TURNO #4"
                End With
            End If
        End With
        With Timer1
            .Enabled = True
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidosPendientes:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Resize()
        On Error GoTo errHandler
        With Frame1(1)
            .Left = 500
            .Top = 250
            .Height = (frmPedidosPendientes.Height - 1500) / 2
            .Width = (frmPedidosPendientes.Width - 1500) / 2
        End With
        
        With Frame1(2)
            .Left = 1000 + Frame1(1).Width
            .Top = Frame1(1).Top
            .Height = Frame1(1).Height
            .Width = Frame1(1).Width
        End With
        
        With Frame1(4)
            .Left = Frame1(1).Left
            .Top = 500 + Frame1(1).Height
            .Height = Frame1(1).Height
            .Width = Frame1(1).Width
        End With
        
        With Frame1(6)
            .Left = Frame1(2).Left
            .Top = Frame1(4).Top
            .Height = Frame1(1).Height
            .Width = Frame1(1).Width
        End With
        
        With Frame1(0)
            .Left = 250
            .Top = 250
            .Height = Frame1(1).Height - 500
            .Width = Frame1(1).Width - 500
        End With
        
        With Frame1(3)
            .Left = Frame1(0).Left
            .Top = Frame1(0).Top
            .Height = Frame1(0).Height
            .Width = Frame1(0).Width
        End With
        
        With Frame1(5)
            .Left = Frame1(0).Left
            .Top = Frame1(0).Top
            .Height = Frame1(0).Height
            .Width = Frame1(0).Width
        End With
        
        With Frame1(7)
            .Left = Frame1(0).Left
            .Top = Frame1(0).Top
            .Height = Frame1(0).Height
            .Width = Frame1(0).Width
        End With
        
        With Label1(0)
            .Left = Frame1(0).Left
            .Top = Frame1(0).Top
            .Height = Frame1(0).Height - 500
            .Width = Frame1(0).Width - 500
            .FontSize = Round(Label1(0).Height / 140)
        End With
        
        With Label1(3)
            .Left = Label1(0).Left
            .Top = Label1(0).Top
            .Height = Label1(0).Height
            .Width = Label1(0).Width
            .FontSize = Label1(0).FontSize
        End With
        
        With Label1(5)
            .Left = Label1(0).Left
            .Top = Label1(0).Top
            .Height = Label1(0).Height
            .Width = Label1(0).Width
            .FontSize = Label1(0).FontSize
        End With
        
        With Label1(7)
            .Left = Label1(0).Left
            .Top = Label1(0).Top
            .Height = Label1(0).Height
            .Width = Label1(0).Width
            .FontSize = Label1(0).FontSize
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidosPendientes:Form_Resize" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Timer1_Timer()
        On Error GoTo errHandler
        With Rs
            If .State = 1 Then .Close
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            .Open "SELECT top 4 min(id) as id,Folio, Nombre From PO_LINES_ALL Where Tipo= 'Pedidos' AND cancelado= 'No' group by Folio, Nombre order by 1;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
            .Requery
            If .RecordCount <> 0 Then
                .MoveFirst
                Label1(0).Caption = "TURNO #1" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                .MoveNext
                If Not .EOF Then
                    Label1(3).Caption = "TURNO #2" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                    .MoveNext
                    If Not .EOF Then
                        Label1(5).Caption = "TURNO #3" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                        .MoveNext
                        If Not .EOF Then
                            Label1(7).Caption = "TURNO #4" & vbNewLine & "FOLIO : " & .Fields(1).Value & vbNewLine & "CLIENTE : " & .Fields(2).Value
                        Else
                            With Label1(7)
                                .Caption = "TURNO #4"
                            End With
                        End If
                    Else
                        With Label1(5)
                            .Caption = "TURNO #3"
                        End With
                        
                        With Label1(7)
                            .Caption = "TURNO #4"
                        End With
                    End If
                Else
                    With Label1(3)
                        .Caption = "TURNO #2"
                    End With
                    
                    With Label1(5)
                        .Caption = "TURNO #3"
                    End With
                    
                    With Label1(7)
                        .Caption = "TURNO #4"
                    End With
                End If
            Else
                With Label1(0)
                    .Caption = "TURNO #1"
                End With
                
                With Label1(3)
                    .Caption = "TURNO #2"
                End With
                
                With Label1(5)
                    .Caption = "TURNO #3"
                End With
                
                With Label1(7)
                    .Caption = "TURNO #4"
                End With
            End If
        End With
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidosPendientes:Timer1_Timer" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
    End Sub

    Private Sub Salir_Click()
        On Error GoTo errHandler
        Unload Me
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidosPendientes:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        With Rs
            If .State = 1 Then .Close
        End With
        
        With Cn
            If .State = 1 Then .Close
        End With
        
        Set Rs = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmPedidosPendientes:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
