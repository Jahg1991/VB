VERSION 5.00
Begin VB.Form frmCrearCorteCaja 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Corte de Caja"
   ClientHeight    =   7470
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   13
         Left            =   1200
         TabIndex        =   29
         Top             =   6480
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   12
         Left            =   1200
         TabIndex        =   28
         Top             =   6000
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   11
         Left            =   1200
         TabIndex        =   27
         Top             =   5520
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   10
         Left            =   1200
         TabIndex        =   26
         Top             =   5040
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   9
         Left            =   1200
         TabIndex        =   25
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   8
         Left            =   1200
         TabIndex        =   24
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   7
         Left            =   1200
         TabIndex        =   23
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   6
         Left            =   1200
         TabIndex        =   22
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   5
         Left            =   1200
         TabIndex        =   21
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   4
         Left            =   1200
         TabIndex        =   20
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   3
         Left            =   1200
         TabIndex        =   19
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   2
         Left            =   1200
         TabIndex        =   18
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   420
         Index           =   0
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4455
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$0.10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   0
            TabIndex        =   15
            Top             =   6360
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$0.20"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   0
            TabIndex        =   14
            Top             =   5880
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$0.50"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   0
            TabIndex        =   13
            Top             =   5400
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   0
            TabIndex        =   12
            Top             =   4920
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   0
            TabIndex        =   11
            Top             =   4440
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   0
            TabIndex        =   10
            Top             =   3960
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$10"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   9
            Top             =   3480
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$20"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   8
            Top             =   3000
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$50"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   7
            Top             =   2520
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$100"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   6
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$200"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   5
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$500"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   4
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$1000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   1005
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Guardar 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmCrearCorteCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    '//RECORDSET
    Dim Rs As New adodb.Recordset   'Corte de Caja
    Dim Rs1 As New adodb.Recordset  'Ticket
    
    '//SUBTOTALES
    Dim T1 As Integer               '1000
    Dim T2 As Integer               '500
    Dim T3 As Integer               '200
    Dim T4 As Integer               '100
    Dim T5 As Integer               '50
    Dim T6 As Integer               '20
    Dim T7 As Integer               '10
    Dim T8 As Integer               '5
    Dim T9 As Integer               '2
    Dim T10 As Integer              '1
    Dim T11 As Integer              '.5
    Dim T12 As Integer              '.2
    Dim T13 As Integer              '.1

    Private Sub Form_Load()
        On Error GoTo errHandler
        
        With Cn
            .CursorLocation = adodb.CursorLocationEnum.adUseClient
            If .State = 0 Then .Open (StConnection)
        End With
    
        T1 = 0
        T2 = 0
        T3 = 0
        T4 = 0
        T5 = 0
        T6 = 0
        T7 = 0
        T8 = 0
        T9 = 0
        T10 = 0
        T11 = 0
        T12 = 0
        T13 = 0
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCrearCorteCaja:Form_Load" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub

    Private Sub Text2_Change(Index As Integer)
        On Error GoTo errHandler
        
        Select Case Index
            Case 1
                If Text2(1) = "" Then
                    T1 = 0
                Else
                    T1 = Val(Text2(1))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 2
                If Text2(2) = "" Then
                    T2 = 0
                Else
                    T2 = Val(Text2(2))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 3
                If Text2(3) = "" Then
                    T3 = 0
                Else
                    T3 = Val(Text2(3))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 4
                If Text2(4) = "" Then
                    T4 = 0
                Else
                    T4 = Val(Text2(4))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 5
                If Text2(5) = "" Then
                    T5 = 0
                Else
                    T5 = Val(Text2(5))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 6
                If Text2(6) = "" Then
                    T6 = 0
                Else
                    T6 = Val(Text2(6))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 7
                If Text2(7) = "" Then
                    T7 = 0
                Else
                    T7 = Val(Text2(7))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 8
                If Text2(8) = "" Then
                    T8 = 0
                Else
                    T8 = Val(Text2(8))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 9
                If Text2(9) = "" Then
                    T9 = 0
                Else
                    T9 = Val(Text2(9))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 10
                If Text2(10) = "" Then
                    T10 = 0
                Else
                    T10 = Val(Text2(10))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 11
                If Text2(11) = "" Then
                    T11 = 0
                Else
                    T11 = Val(Text2(11))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 12
                If Text2(12) = "" Then
                    T12 = 0
                Else
                    T12 = Val(Text2(12))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
                
            Case 13
                If Text2(13) = "" Then
                    T13 = 0
                Else
                    T13 = Val(Text2(13))
                End If
                
                Text2(0) = Format(((Val(T1) * 1000) + (Val(T2) * 500) + (Val(T3) * 200) + (Val(T4) * 100) + (Val(T5) * 50) + (Val(T6) * 20) + (Val(T7) * 10) + (Val(T8) * 5) + (Val(T9) * 2) + (Val(T10) * 1) + (Val(T11) * 0.5) + (Val(T12) * 0.2) + (Val(T13) * 0.1)), "0.00")
        End Select
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCrearCorteCaja:Text2_Change" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
        Private Sub Guardar_Click()
        On Error GoTo errHandler
        
        vbq = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo, "Información")
                    
        If vbq = vbYes Then
            If Val(Text2(0)) = 0 Or Text2(0) = "" Then
                MsgBox "No hay información que guardar", vbOKOnly, "Advertencia"
                
                Exit Sub
            End If
            
            With Rs
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select * from CE_BOX_CUT;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
                .AddNew
                    .Fields(1) = Date
                    If Text2(0) <> "" Then .Fields(2) = Replace(Text2(0), ",", ".")
                    If Text2(1) <> "" Then .Fields(3) = Round(Val(Text2(1)))
                    If Text2(2) <> "" Then .Fields(4) = Round(Val(Text2(2)))
                    If Text2(3) <> "" Then .Fields(5) = Round(Val(Text2(3)))
                    If Text2(4) <> "" Then .Fields(6) = Round(Val(Text2(4)))
                    If Text2(5) <> "" Then .Fields(7) = Round(Val(Text2(5)))
                    If Text2(6) <> "" Then .Fields(8) = Round(Val(Text2(6)))
                    If Text2(7) <> "" Then .Fields(9) = Round(Val(Text2(7)))
                    If Text2(8) <> "" Then .Fields(10) = Round(Val(Text2(8)))
                    If Text2(9) <> "" Then .Fields(11) = Round(Val(Text2(9)))
                    If Text2(10) <> "" Then .Fields(12) = Round(Val(Text2(10)))
                    If Text2(11) <> "" Then .Fields(13) = Round(Val(Text2(11)))
                    If Text2(12) <> "" Then .Fields(14) = Round(Val(Text2(12)))
                    If Text2(13) <> "" Then .Fields(15) = Round(Val(Text2(13)))
                    .Fields(16) = frmMenuInicial.Combo1.Text
                    .Fields(17) = StUsuario
                .Update
                .Requery
                .Close
            End With
            
            MsgBox "Corte de caja terminado", vbOKOnly, "Finalizado"
            
            With Rs1
                If .State = 1 Then .Close
                .CursorLocation = adodb.CursorLocationEnum.adUseClient
                .Open "Select top 1 * from CE_BOX_CUT where caja = '" & frmMenuInicial.Combo1.Text & "' order by id desc;", Cn, adodb.CursorTypeEnum.adOpenStatic, adodb.LockTypeEnum.adLockOptimistic
                .Requery
            End With
            
            With TicketCorte2
                Set .DataSource = Rs1
            
                With .Sections("Sección4")
                    .Controls("Etiqueta2").Caption = frmMenuInicial.Combo1.Text
                    If IsNull(Rs1.Fields(3).Value) = False Then .Controls("Etiqueta10").Caption = Rs1.Fields(3).Value
                    If IsNull(Rs1.Fields(4).Value) = False Then .Controls("Etiqueta11").Caption = Rs1.Fields(4).Value
                    If IsNull(Rs1.Fields(5).Value) = False Then .Controls("Etiqueta12").Caption = Rs1.Fields(5).Value
                    If IsNull(Rs1.Fields(6).Value) = False Then .Controls("Etiqueta13").Caption = Rs1.Fields(6).Value
                    If IsNull(Rs1.Fields(7).Value) = False Then .Controls("Etiqueta14").Caption = Rs1.Fields(7).Value
                    If IsNull(Rs1.Fields(8).Value) = False Then .Controls("Etiqueta15").Caption = Rs1.Fields(8).Value
                    If IsNull(Rs1.Fields(9).Value) = False Then .Controls("Etiqueta26").Caption = Rs1.Fields(9).Value
                    If IsNull(Rs1.Fields(10).Value) = False Then .Controls("Etiqueta27").Caption = Rs1.Fields(10).Value
                    If IsNull(Rs1.Fields(11).Value) = False Then .Controls("Etiqueta28").Caption = Rs1.Fields(11).Value
                    If IsNull(Rs1.Fields(12).Value) = False Then .Controls("Etiqueta29").Caption = Rs1.Fields(12).Value
                    If IsNull(Rs1.Fields(13).Value) = False Then .Controls("Etiqueta30").Caption = Rs1.Fields(13).Value
                    If IsNull(Rs1.Fields(14).Value) = False Then .Controls("Etiqueta31").Caption = Rs1.Fields(14).Value
                    If IsNull(Rs1.Fields(15).Value) = False Then .Controls("Etiqueta32").Caption = Rs1.Fields(15).Value
                    .Controls("Etiqueta33").Caption = Rs1.Fields(2).Value
                End With
                
                .Show 1
            End With
            
            Rs1.Close
    
            Unload Me
        Else
            Exit Sub
        End If
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCrearCorteCaja:Guardar_Click" & vbTab & err.Number & vbTab & err.Description
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
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCrearCorteCaja:Salir_Click" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error GoTo errHandler
        
        If Rs.State = 1 Then Rs.Close
        If Rs1.State = 1 Then Rs.Close
        If Cn.State = 1 Then Cn.Close
        
        Set Rs = Nothing
        Set Rs1 = Nothing
        Set Cn = Nothing
    Exit Sub
errHandler:
        FileNum = FreeFile
        Open App.Path & "\ErrorRegistry.txt" For Append As FileNum
        Print #FileNum, Format(Date, "YYYY-MM-DD") & vbTab & Format(Time, "HH:MM:SS") & vbTab & "Error en: frmCrearCorteCaja:Form_Unload" & vbTab & err.Number & vbTab & err.Description
        Close FileNum
        err.Clear
        
        MsgBox "Hubo un error consulte la bitacora", vbInformation, "Error"
    End Sub
