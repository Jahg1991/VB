VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14055
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13815
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Detalles"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2655
            Left            =   120
            TabIndex        =   3
            Top             =   3000
            Width           =   13575
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Consolas"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2280
               Index           =   1
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   13335
            End
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2280
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   13575
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   5
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   10404034
            CalendarTrailingForeColor=   10404034
            Format          =   115474433
            CurrentDate     =   43466
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Index           =   1
            Left            =   3240
            TabIndex        =   6
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   10404034
            CalendarTrailingForeColor=   10404034
            Format          =   115474433
            CurrentDate     =   43812
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   7
            Top             =   120
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DTPicker1_Change(Index As Integer)
    
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim c5 As String
    Dim nc1 As Integer
    Dim nc2 As Integer
    Dim nc3 As Integer
    Dim nc4 As Integer
    Dim nc5 As Integer
    Select Case Index
        Case 0
            If TipoCatalogo = 2 Then
                List1(0).Clear
                RsPagosV.Requery
                RsPagosV.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                Do Until RsPagosV.EOF
                    'JAHG Longitud Cadena
                    c1 = Mid(RsPagosV!Fecha, 1, 10)
                    c2 = Mid(RsPagosV!Folio, 1, 15)
                    c3 = Mid(RsPagosV!Nombre, 1, 45)
                    c4 = Mid(RsPagosV!Referencia, 1, 25)
                    c5 = Mid(RsPagosV!Total, 1, 20)
                    nc1 = 10 - Len(c1)
                    nc2 = 15 - Len(c2)
                    nc3 = 45 - Len(c3)
                    nc4 = 25 - Len(c4)
                    nc5 = 20 - Len(c5)
                    For i = 1 To nc1
                        c1 = c1 & " "
                    Next i
                    For i = 1 To nc2
                        c2 = c2 & " "
                    Next i
                    For i = 1 To nc3
                        c3 = c3 & " "
                    Next i
                    For i = 1 To nc4
                        c4 = c4 & " "
                    Next i
                    For i = 1 To nc5
                        c5 = " " & c5
                    Next i
                    List1(0).AddItem c1 & " |" & c2 & " |" & c3 & " |" & c4 & " |$" & c5
                    RsPagosV.MoveNext
                Loop
            End If
            If TipoCatalogo = 3 Then
                List1(0).Clear
                List1(1).Clear
                RsCabeceraVentas.Requery
                RsCabeceraVentas.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                Do Until RsCabeceraVentas.EOF
                    'JAHG Longitud Cadena
                    c1 = Mid(RsCabeceraVentas!Fecha, 1, 10)
                    c2 = Mid(RsCabeceraVentas!Folio, 1, 15)
                    c3 = Mid(RsCabeceraVentas!Nombre, 1, 70)
                    c4 = Mid(RsCabeceraVentas!Total, 1, 20)
                    nc1 = 10 - Len(c1)
                    nc2 = 15 - Len(c2)
                    nc3 = 70 - Len(c3)
                    nc4 = 20 - Len(c4)
                    For i = 1 To nc1
                        c1 = c1 & " "
                    Next i
                    For i = 1 To nc2
                        c2 = c2 & " "
                    Next i
                    For i = 1 To nc3
                        c3 = c3 & " "
                    Next i
                    For i = 1 To nc4
                        c4 = " " & c4
                    Next i
                    List1(0).AddItem c1 & " |" & c2 & " |" & c3 & " |$" & c4
                    RsCabeceraVentas.MoveNext
                Loop
            End If
        Case 1
            If TipoCatalogo = 2 Then
                List1(0).Clear
                RsPagosV.Requery
                RsPagosV.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                Do Until RsPagosV.EOF
                    'JAHG Longitud Cadena
                    c1 = Mid(RsPagosV!Fecha, 1, 10)
                    c2 = Mid(RsPagosV!Folio, 1, 15)
                    c3 = Mid(RsPagosV!Nombre, 1, 45)
                    c4 = Mid(RsPagosV!Referencia, 1, 25)
                    c5 = Mid(RsPagosV!Total, 1, 20)
                    nc1 = 10 - Len(c1)
                    nc2 = 15 - Len(c2)
                    nc3 = 45 - Len(c3)
                    nc4 = 25 - Len(c4)
                    nc5 = 20 - Len(c5)
                    For i = 1 To nc1
                        c1 = c1 & " "
                    Next i
                    For i = 1 To nc2
                        c2 = c2 & " "
                    Next i
                    For i = 1 To nc3
                        c3 = c3 & " "
                    Next i
                    For i = 1 To nc4
                        c4 = c4 & " "
                    Next i
                    For i = 1 To nc5
                        c5 = " " & c5
                    Next i
                    List1(0).AddItem c1 & " |" & c2 & " |" & c3 & " |" & c4 & " |$" & c5
                    RsPagosV.MoveNext
                Loop
            End If
            If TipoCatalogo = 3 Then
                List1(0).Clear
                List1(1).Clear
                RsCabeceraVentas.Requery
                RsCabeceraVentas.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
                Do Until RsCabeceraVentas.EOF
                    'JAHG Longitud Cadena
                    c1 = Mid(RsCabeceraVentas!Fecha, 1, 10)
                    c2 = Mid(RsCabeceraVentas!Folio, 1, 15)
                    c3 = Mid(RsCabeceraVentas!Nombre, 1, 70)
                    c4 = Mid(RsCabeceraVentas!Total, 1, 20)
                    nc1 = 10 - Len(c1)
                    nc2 = 15 - Len(c2)
                    nc3 = 70 - Len(c3)
                    nc4 = 20 - Len(c4)
                    For i = 1 To nc1
                        c1 = c1 & " "
                    Next i
                    For i = 1 To nc2
                        c2 = c2 & " "
                    Next i
                    For i = 1 To nc3
                        c3 = c3 & " "
                    Next i
                    For i = 1 To nc4
                        c4 = " " & c4
                    Next i
                    List1(0).AddItem c1 & " |" & c2 & " |" & c3 & " |$" & c4
                    RsCabeceraVentas.MoveNext
                Loop
            End If
    End Select
    
End Sub

Private Sub Form_Load()

    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim c5 As String
    Dim nc1 As Integer
    Dim nc2 As Integer
    Dim nc3 As Integer
    Dim nc4 As Integer
    Dim nc5 As Integer
    DTPicker1(0).Value = Date
    DTPicker1(1).Value = Date
    If TipoCatalogo = 2 Then
        Form5.Icon = LoadPicture(App.Path & "\Images\Control de pagos.ico")
        Form5.Caption = "Control de Pagos"
        Frame2.Visible = False
        List1(0).Height = 4980
        List1(0).Clear
        RsPagosV.Requery
        RsPagosV.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
        Do Until RsPagosV.EOF
            'JAHG Longitud Cadena
            c1 = Mid(RsPagosV!Fecha, 1, 10)
            c2 = Mid(RsPagosV!Folio, 1, 15)
            c3 = Mid(RsPagosV!Nombre, 1, 45)
            c4 = Mid(RsPagosV!Referencia, 1, 25)
            c5 = Mid(RsPagosV!Total, 1, 20)
            nc1 = 10 - Len(c1)
            nc2 = 15 - Len(c2)
            nc3 = 45 - Len(c3)
            nc4 = 25 - Len(c4)
            nc5 = 20 - Len(c5)
            For i = 1 To nc1
                c1 = c1 & " "
            Next i
            For i = 1 To nc2
                c2 = c2 & " "
            Next i
            For i = 1 To nc3
                c3 = c3 & " "
            Next i
            For i = 1 To nc4
                c4 = c4 & " "
            Next i
            For i = 1 To nc5
                c5 = " " & c5
            Next i
            List1(0).AddItem c1 & " |" & c2 & " |" & c3 & " |" & c4 & " |$" & c5
            RsPagosV.MoveNext
        Loop
    End If
    If TipoCatalogo = 3 Then
        Form5.Icon = LoadPicture(App.Path & "\Images\Control de ventas.ico")
        Form5.Caption = "Control de Ventas"
        List1(0).Clear
        List1(1).Clear
        RsCabeceraVentas.Requery
        RsCabeceraVentas.Filter = "Fecha >= '" & DTPicker1(0).Value & "' and  Fecha <= '" & DTPicker1(1).Value & "' "
        Do Until RsCabeceraVentas.EOF
            'JAHG Longitud Cadena
            c1 = Mid(RsCabeceraVentas!Fecha, 1, 10)
            c2 = Mid(RsCabeceraVentas!Folio, 1, 15)
            c3 = Mid(RsCabeceraVentas!Nombre, 1, 70)
            c4 = Mid(RsCabeceraVentas!Total, 1, 20)
            nc1 = 10 - Len(c1)
            nc2 = 15 - Len(c2)
            nc3 = 70 - Len(c3)
            nc4 = 20 - Len(c4)
            For i = 1 To nc1
                c1 = c1 & " "
            Next i
            For i = 1 To nc2
                c2 = c2 & " "
            Next i
            For i = 1 To nc3
                c3 = c3 & " "
            Next i
            For i = 1 To nc4
                c4 = " " & c4
            Next i
            List1(0).AddItem c1 & " |" & c2 & " |" & c3 & " |$" & c4
            RsCabeceraVentas.MoveNext
        Loop
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If TipoCatalogo = 2 Then
        RsPagosV.Filter = ""
    End If
    If TipoCatalogo = 3 Then
        RsCabeceraVentas.Filter = ""
        RsVentasV.Filter = ""
    End If
    Form1.Enabled = True

End Sub

Private Sub List1_Click(Index As Integer)
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim c5 As String
    Dim c6 As String
    Dim nc1 As Integer
    Dim nc2 As Integer
    Dim nc3 As Integer
    Dim nc4 As Integer
    Dim nc5 As Integer
    Dim nc6 As String
    Select Case Index
        Case 0
            If TipoCatalogo = 3 Then
                List1(1).Clear
                RsVentasV.Requery
                RsVentasV.Filter = "Folio = '" & Trim(Mid(List1(0).Text, 13, 15)) & "'"
                Do Until RsVentasV.EOF
                    'JAHG Longitud Cadena
                    c1 = Mid(RsVentasV!Codigo, 1, 25)
                    c2 = Mid(RsVentasV!Descripcion, 1, 40)
                    c3 = Mid(RsVentasV!Cantidad, 1, 15)
                    c4 = Mid(RsVentasV!UDM, 1, 2)
                    c5 = Mid(RsVentasV!Precio, 1, 7)
                    c6 = Mid(RsVentasV!Total, 1, 20)
                    nc1 = 25 - Len(c1)
                    nc2 = 40 - Len(c2)
                    nc3 = 15 - Len(c3)
                    nc4 = 2 - Len(c4)
                    nc5 = 7 - Len(c5)
                    nc6 = 20 - Len(c6)
                    For i = 1 To nc1
                        c1 = c1 & " "
                    Next i
                    For i = 1 To nc2
                        c2 = c2 & " "
                    Next i
                    For i = 1 To nc3
                        c3 = " " & c3
                    Next i
                    For i = 1 To nc4
                        c4 = c4 & " "
                    Next i
                    For i = 1 To nc5
                        c5 = " " & c5
                    Next i
                    For i = 1 To nc6
                        c6 = " " & c6
                    Next i
                    List1(1).AddItem c1 & " |" & c2 & " |" & c3 & " |" & c4 & " |$" & c5 & " |$" & c6
                    RsVentasV.MoveNext
                Loop
            End If
    End Select
    
End Sub
