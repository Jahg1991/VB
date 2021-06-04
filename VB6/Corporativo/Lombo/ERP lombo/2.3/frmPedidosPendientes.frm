VERSION 5.00
Begin VB.Form frmPedidosPendientes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Pendientes"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   13890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   13890
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6840
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   6
      Left            =   7080
      TabIndex        =   9
      Top             =   2760
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
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
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
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
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
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
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   2
      Left            =   7080
      TabIndex        =   3
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
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
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
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
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
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
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
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

    Option Explicit
    
    '//RECORDSET
    Dim Rs                  As New adodb.Recordset
    
    Private Sub Form_Load()
        On Error GoTo errHandler
        
        frmPedidosPendientes.WindowState = 2
        
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
                Label1(0).Caption = "Turno #1" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                .MoveNext
                If Not .EOF Then
                    Label1(3).Caption = "Turno #2" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                    .MoveNext
                    If Not .EOF Then
                        Label1(5).Caption = "Turno #3" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                        .MoveNext
                        If Not .EOF Then
                            Label1(7).Caption = "Turno #4" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                        Else
                            Label1(7).Caption = "Turno #4"
                        End If
                    Else
                        Label1(5).Caption = "Turno #3"
                        Label1(7).Caption = "Turno #4"
                    End If
                Else
                    Label1(3).Caption = "Turno #2"
                    Label1(5).Caption = "Turno #3"
                    Label1(7).Caption = "Turno #4"
                End If
            Else
                Label1(0).Caption = "Turno #1"
                Label1(3).Caption = "Turno #2"
                Label1(5).Caption = "Turno #3"
                Label1(7).Caption = "Turno #4"
            End If
        End With
        
        Timer1.Enabled = True
        
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
        
        Frame1(1).Left = 500
        Frame1(1).Top = 250
        Frame1(1).Height = (frmPedidosPendientes.Height - 1500) / 2
        Frame1(1).Width = (frmPedidosPendientes.Width - 1500) / 2
        
        Frame1(2).Left = 1000 + Frame1(1).Width
        Frame1(2).Top = Frame1(1).Top
        Frame1(2).Height = Frame1(1).Height
        Frame1(2).Width = Frame1(1).Width
        
        Frame1(4).Left = Frame1(1).Left
        Frame1(4).Top = 500 + Frame1(1).Height
        Frame1(4).Height = Frame1(1).Height
        Frame1(4).Width = Frame1(1).Width
        
        Frame1(6).Left = Frame1(2).Left
        Frame1(6).Top = Frame1(4).Top
        Frame1(6).Height = Frame1(1).Height
        Frame1(6).Width = Frame1(1).Width
        
        Frame1(0).Left = 250
        Frame1(0).Top = 250
        Frame1(0).Height = Frame1(1).Height - 500
        Frame1(0).Width = Frame1(1).Width - 500
        
        Frame1(3).Left = Frame1(0).Left
        Frame1(3).Top = Frame1(0).Top
        Frame1(3).Height = Frame1(0).Height
        Frame1(3).Width = Frame1(0).Width
        
        Frame1(5).Left = Frame1(0).Left
        Frame1(5).Top = Frame1(0).Top
        Frame1(5).Height = Frame1(0).Height
        Frame1(5).Width = Frame1(0).Width
        
        Frame1(7).Left = Frame1(0).Left
        Frame1(7).Top = Frame1(0).Top
        Frame1(7).Height = Frame1(0).Height
        Frame1(7).Width = Frame1(0).Width
        
        Label1(0).Left = Frame1(0).Left
        Label1(0).Top = Frame1(0).Top
        Label1(0).Height = Frame1(0).Height - 500
        Label1(0).Width = Frame1(0).Width - 500
        
        Label1(3).Left = Label1(0).Left
        Label1(3).Top = Label1(0).Top
        Label1(3).Height = Label1(0).Height
        Label1(3).Width = Label1(0).Width
        
        Label1(5).Left = Label1(0).Left
        Label1(5).Top = Label1(0).Top
        Label1(5).Height = Label1(0).Height
        Label1(5).Width = Label1(0).Width
        
        Label1(7).Left = Label1(0).Left
        Label1(7).Top = Label1(0).Top
        Label1(7).Height = Label1(0).Height
        Label1(7).Width = Label1(0).Width
        
        Label1(0).FontSize = Round(Label1(0).Height / 100)
        Label1(3).FontSize = Label1(0).FontSize
        Label1(5).FontSize = Label1(0).FontSize
        Label1(7).FontSize = Label1(0).FontSize
        
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
            .Requery
            If .RecordCount <> 0 Then
                .MoveFirst
                Label1(0).Caption = "Turno #1" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                .MoveNext
                If Not .EOF Then
                    Label1(3).Caption = "Turno #2" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                    .MoveNext
                    If Not .EOF Then
                        Label1(5).Caption = "Turno #3" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                        .MoveNext
                        If Not .EOF Then
                            Label1(7).Caption = "Turno #4" & vbNewLine & "Folio : " & Rs.Fields(1).Value & vbNewLine & "Cliente : " & Rs.Fields(2).Value
                        Else
                            Label1(7).Caption = "Turno #4"
                        End If
                    Else
                        Label1(5).Caption = "Turno #3"
                        Label1(7).Caption = "Turno #4"
                    End If
                Else
                    Label1(3).Caption = "Turno #2"
                    Label1(5).Caption = "Turno #3"
                    Label1(7).Caption = "Turno #4"
                End If
            Else
                Label1(0).Caption = "Turno #1"
                Label1(3).Caption = "Turno #2"
                Label1(5).Caption = "Turno #3"
                Label1(7).Caption = "Turno #4"
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
        
        If Rs.State = 1 Then Rs.Close
        If Cn.State = 1 Then Cn.Close
        
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
