VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
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
   ScaleHeight     =   5325
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   5480
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   5000
         Begin VB.TextBox Text5 
            Height          =   420
            Left            =   1800
            TabIndex        =   37
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Height          =   420
            Left            =   1800
            TabIndex        =   34
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Imprimir ticket y cambiar de cliente"
            Height          =   375
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3360
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   375
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   3360
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Height          =   420
            Left            =   1800
            TabIndex        =   30
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            Height          =   420
            Index           =   2
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1800
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            Height          =   420
            Index           =   1
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            Height          =   420
            Index           =   0
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "N�mero de lote"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   36
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Cliente"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Kilogramos"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Producto"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Lugar de destino"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Lugar de origen"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   5000
         Begin VB.TextBox Text2 
            Height          =   420
            Index           =   1
            Left            =   1680
            TabIndex        =   21
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   420
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   420
            Index           =   0
            Left            =   1680
            TabIndex        =   20
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            Height          =   420
            Index           =   1
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   840
            Width           =   1815
         End
         Begin VB.ComboBox Combo2 
            Height          =   420
            Index           =   0
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "N�mero de lote"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   35
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Kilogramos"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Producto"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Lugar de origen"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5000
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Guardar"
            Height          =   495
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   420
            Index           =   1
            Left            =   2520
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   420
            Index           =   0
            Left            =   2520
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            Height          =   420
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   2200
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Kilogramos"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   2200
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad de Capotes"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   2145
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Lugar de origen"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Sub ValidarDirectorios()
    
    On Error GoTo ErrorDirectorio
    
    i = GetAttr(App.Path & "\ERP")
    
    Exit Sub
ErrorDirectorio:
        If Err.Number = 53 Then
            MkDir App.Path & "\ERP"
        End If
    End Sub


Private Sub Form_Load()
    With Form1
        .BackColor = &HFFFFFF
        .WindowState = 2
        
        With Command1(0)
            .BackColor = &HC0C0C0
            .Caption = "Conversi�n de piezas a Kg"
        End With
        
        With Command1(1)
            .BackColor = &HC0C0C0
            .Caption = "Conversi�n a PT"
        End With
        
        With Command1(2)
            .BackColor = &HC0C0C0
            .Caption = "Traspasos de PT"
        End With
        
        With Command1(3)
            .BackColor = &HC0C0C0
            .Caption = "Salir"
        End With
        
        With FramePrincipal
            .BackColor = &HC0C0C0
            .Caption = ""
        End With
        
        With Frame1(0)
            .BackColor = &H8000000F
            .Visible = False
            .Caption = "Conversi�n de piezas a Kg"
        End With
        
        With Frame1(1)
            .BackColor = &H8000000F
            .Visible = False
            .Caption = "Conversi�n a PT"
        End With
        
        With Frame1(2)
            .BackColor = &H8000000F
            .Visible = False
            .Caption = "Traspasos de PT"
        End With
        
        With Image1
            .Visible = True
            .Picture = LoadPicture(App.Path & "\Imagenes\Inicio.jpg")
        End With
    End With
End Sub

Private Sub Form_Resize()
    With Form1
        For i = 0 To 3
            With Command1(i)
                .Height = Round(Form1.Height / 5)
                .Width = Round(Form1.Width / 3)
                .Left = Round(Form1.Width / 27)
                .FontSize = Round(Form1.Height / 250)
            End With
        Next i
        
        With Command1(0)
            .Top = Round(Form1.Height / 25)
        End With
        
        With Command1(1)
            .Top = Command1(0).Top + Command1(0).Height + Round(Form1.Height / 25)
        End With
        
        With Command1(2)
            .Top = Command1(1).Top + Command1(1).Height + Round(Form1.Height / 25)
        End With
        
        With Command1(3)
            .Top = Command1(2).Top + Command1(2).Height + Round(Form1.Height / 25)
        End With
        
        With FramePrincipal
            .Height = Command1(2).Top + Command1(2).Height * 2
            .Width = Round(Form1.Width / 1.8)
            .Left = Round(Form1.Width / 2.5)
            .Top = Round(Form1.Height / 5 / 5)
            .FontSize = Round(Form1.Height / 250)
        End With
        
        For i = 0 To 2
            With Frame1(i)
                .Height = FramePrincipal.Height - 480
                .Width = FramePrincipal.Width - 480
                .Left = 240
                .Top = 240
                .FontSize = Round(Form1.Height / 500)
            End With
        Next i
        
        With Image1
            .Height = FramePrincipal.Height - 480
            .Width = FramePrincipal.Width - 480
            .Left = 240
            .Top = 240
        End With
        
        For i = 0 To 2
            With Label1(i)
                .Width = Round(Frame1(0).Width / 3)
                .Height = Round(.Width / 5.6)
                .Left = 240
                .FontSize = Round(.Height / 32)
            End With
        Next i
        
        With Label1(0)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label1(1)
            .Top = Round(.Height * 3)
        End With
        
        With Label1(2)
            .Top = Round(.Height * 4.5)
        End With
        
        With Combo1
            .Width = Round(Frame1(0).Width / 1.8)
            .Left = Round(Frame1(0).Width / 3) + 480
            .Top = Label1(0).Top
            .FontSize = Label1(0).FontSize
        End With
        
        For i = 0 To 1
            With Text1(i)
                .Width = Combo1.Width
                .Height = Combo1.Height
                .Left = Combo1.Left
                .FontSize = Combo1.FontSize
            End With
        Next i

        With Text1(0)
            .Top = Label1(1).Top
        End With
        
        With Text1(1)
            .Top = Label1(2).Top
        End With
        
        With Command2
            .Width = Frame1(0).Width / 2
            .Height = Label1(2).Height * 1.5
            .Top = Label1(2).Top + Label1(2).Height * 2
            .Left = Frame1(0).Width / 4
            .FontSize = Label1(2).FontSize
        End With
    
        For i = 0 To 3
            With Label2(i)
                .Width = Round(Frame1(1).Width / 3)
                .Height = Round(.Width / 5.6)
                .Left = 240
                .FontSize = Round(.Height / 32)
            End With
        Next i
        
        With Label2(0)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label2(1)
            .Top = Round(.Height * 3)
        End With
        
        With Label2(2)
            .Top = Round(.Height * 4.5)
        End With
        
        With Label2(3)
            .Top = Round(.Height * 6)
        End With
        
        For i = 0 To 1
            With Combo2(i)
                .Width = Round(Frame1(1).Width / 1.8)
                .Left = Round(Frame1(1).Width / 3) + 480
                .Top = Label2(i).Top
                .FontSize = Label2(i).FontSize
            End With
        Next i
        
        For i = 0 To 1
            With Text2(i)
                .Width = Combo2(0).Width
                .Height = Combo2(0).Height
                .Left = Combo2(0).Left
                .FontSize = Combo2(0).FontSize
            End With
        Next i
        
        With Text2(0)
            .Top = Label2(2).Top
        End With
        
        With Text2(1)
            .Top = Label2(3).Top
        End With
        
        With Command3
            .Width = Frame1(1).Width / 2
            .Height = Label2(2).Height * 1.5
            .Top = Label2(3).Top + Label2(3).Height * 2
            .Left = Frame1(1).Width / 4
            .FontSize = Label2(3).FontSize
        End With
        
        For i = 0 To 5
            With Label3(i)
                .Width = Round(Frame1(2).Width / 3)
                .Height = Round(.Width / 5.6)
                .Left = 240
                .FontSize = Round(.Height / 32)
            End With
        Next i
        
        With Label3(0)
            .Top = Round(.Height * 1.5)
        End With
        
        With Label3(1)
            .Top = Round(.Height * 3)
        End With
        
        With Label3(2)
            .Top = Round(.Height * 6)
        End With
        
        With Label3(3)
            .Top = Round(.Height * 7.5)
        End With
        
        With Label3(4)
            .Top = Round(.Height * 4.5)
        End With
        
        With Label3(5)
            .Top = Round(.Height * 9)
        End With
        
        For i = 0 To 2
            With Combo3(i)
                .Width = Round(Frame1(2).Width / 1.8)
                .Left = Round(Frame1(2).Width / 3) + 480
                .Top = Label3(i).Top
                .FontSize = Label3(i).FontSize
            End With
        Next i
        
        With Text3
            .Width = Combo3(0).Width
            .Height = Combo3(0).Height
            .Left = Combo3(0).Left
            .Top = Label3(3).Top
            .FontSize = Combo3(0).FontSize
        End With
        
        With Text4
            .Width = Combo3(0).Width
            .Height = Combo3(0).Height
            .Left = Combo3(0).Left
            .Top = Label3(4).Top
            .FontSize = Combo3(0).FontSize
        End With
        
        With Text5
            .Width = Combo3(0).Width
            .Height = Combo3(0).Height
            .Left = Combo3(0).Left
            .Top = Label3(5).Top
            .FontSize = Combo3(0).FontSize
        End With
        
        With Command4
            .Width = Frame1(2).Width / 2
            .Height = Label3(3).Height * 1.5
            .Top = Label3(5).Top + Label3(5).Height * 2
            .Left = Frame1(2).Width / 4
            .FontSize = Label3(3).FontSize
        End With
        
        With Command5
            .Width = Frame1(2).Width / 1.5
            .Height = Label3(3).Height * 1.5
            .Top = Command4.Top + Label3(5).Height * 2
            .Left = Frame1(2).Width / 6
            .FontSize = Label3(3).FontSize
        End With
    End With
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Frame1(0).Visible = True Then
                MsgBox "La pantalla ya est� abierta", vbOKOnly, "Informaci�n"
                Exit Sub
            End If
            
            If Frame1(1).Visible = True Or Frame1(2).Visible = True Then
                vbq = MsgBox("�Desea cerrar la pantalla abierta para abrir la pantalla Conversi�n de piezas a Kg?", vbQuestion + vbYesNo, "Advertencia")
                
                If vbq = vbNo Then
                    Exit Sub
                End If
            End If
                
            With Frame1(0)
                .Visible = True
            End With
                
            With Frame1(1)
                .Visible = False
            End With
                
            With Frame1(2)
                .Visible = False
            End With
            
            With Image1
                .Visible = False
            End With
            
            With Combo1
                .Clear
            End With
            
            For i = 0 To 1
                With Text1(i)
                    .Text = ""
                End With
            Next i
            
            Dim CN As New ADODB.Connection
            Dim RS As New ADODB.Recordset
            
            With CN
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
            End With
                
            With RS
                .CursorLocation = adUseClient
                .Open "Select * from HR_ALL_ORGANIZATION_UNITS Order by NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo1.AddItem .Fields(1).Value & " [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With CN
                .Close
            End With
        Case 1
            If Frame1(1).Visible = True Then
                MsgBox "La pantalla ya est� abierta", vbOKOnly, "Informaci�n"
                Exit Sub
            End If
            
            If Frame1(0).Visible = True Or Frame1(2).Visible = True Then
                vbq = MsgBox("�Desea cerrar la pantalla abierta para abrir la pantalla Conversi�n a PT?", vbQuestion + vbYesNo, "Advertencia")
                
                If vbq = vbNo Then
                    Exit Sub
                End If
            End If
                
            With Frame1(0)
                .Visible = False
            End With
                
            With Frame1(1)
                .Visible = True
            End With
                
            With Frame1(2)
                .Visible = False
            End With
            
            With Image1
                .Visible = False
            End With
            
            For i = 0 To 1
                With Combo2(i)
                    .Clear
                End With
            Next i
            
            For i = 0 To 1
                With Text2(i)
                    .Text = ""
                End With
            Next i
            
            With CN
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
            End With
                
            With RS
                .CursorLocation = adUseClient
                .Open "Select * from HR_ALL_ORGANIZATION_UNITS Order by NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo2(0).AddItem .Fields(1).Value & " [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With RS
                .CursorLocation = adUseClient
                .Open "Select * from MTL_SYSTEM_ITEMS_B Where ENABLED_FLAG = 'Y' Order by DESCRIPTION", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo2(1).AddItem .Fields(2).Value & " [" & .Fields(1).Value & "] [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With CN
                .Close
            End With
        Case 2
            If Frame1(2).Visible = True Then
                MsgBox "La pantalla ya est� abierta", vbOKOnly, "Informaci�n"
                Exit Sub
            End If
            
            If Frame1(0).Visible = True Or Frame1(1).Visible = True Then
                vbq = MsgBox("�Desea cerrar la pantalla abierta para abrir la pantalla Traspasos de PT?", vbQuestion + vbYesNo, "Advertencia")
                
                If vbq = vbNo Then
                    Exit Sub
                End If
            End If
                
            With Frame1(0)
                .Visible = False
            End With
                
            With Frame1(1)
                .Visible = False
            End With
                
            With Frame1(2)
                .Visible = True
            End With
            
            With Image1
                .Visible = False
            End With
            
            For i = 0 To 2
                With Combo3(i)
                    .Clear
                End With
            Next i
            
            With Text3
                .Text = ""
            End With
            
            With Text4
                .Text = ""
            End With
            
            With Text5
                .Text = ""
            End With
            
            With CN
                .CursorLocation = adUseClient
                .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
            End With
                
            With RS
                .CursorLocation = adUseClient
                .Open "Select * from HR_ALL_ORGANIZATION_UNITS Order by NAME", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo3(0).AddItem .Fields(1).Value & " [" & .Fields(0).Value & "]"
                        
                        Combo3(1).AddItem .Fields(1).Value & " [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With RS
                .CursorLocation = adUseClient
                .Open "Select * from MTL_SYSTEM_ITEMS_B Where ENABLED_FLAG = 'Y'  Order by DESCRIPTION", CN, adOpenStatic, adLockOptimistic
                
                If .RecordCount <> 0 Then
                    .MoveFirst
                    
                    While Not .EOF
                        Combo3(2).AddItem .Fields(2).Value & " [" & .Fields(1).Value & "] [" & .Fields(0).Value & "]"
                        
                        .MoveNext
                    Wend
                End If
                
                .Close
            End With
            
            With CN
                .Close
            End With
        Case 3
            vbq = MsgBox("�Desea cerrar el programa?", vbQuestion + vbYesNo, "Advertencia")
                    
            If vbq = vbYes Then
                ValidarDirectorios
                
                With CN
                    .CursorLocation = adUseClient
                    .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
                End With
                    
                With RS
                    .CursorLocation = adUseClient
                    .Open "Select * from WIP_DISCRETE_JOBS Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\WDJ_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                        
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|"
                            
                            .Fields(8).Value = 1
                            .Update
                            .MoveNext
                        Wend
                        
                        Close #1
                    End If
                    
                    .Close
                End With
                
                With RS
                    .CursorLocation = adUseClient
                    .Open "Select * from MTL_MATERIAL_TRANSACTIONS Where TRANSFER_STATUS = 0", CN, adOpenStatic, adLockOptimistic
                    
                    If .RecordCount <> 0 Then
                        Open App.Path & "\ERP\MMT_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1
                        
                        .MoveFirst
                        
                        While Not .EOF
                            Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|" & .Fields(7).Value & "|" & .Fields(8).Value & "|" & .Fields(9).Value & "|" & .Fields(10).Value & "|" & .Fields(11).Value & "|" & .Fields(12).Value & "|" & .Fields(13).Value & "|" & .Fields(14).Value & "|" & .Fields(15).Value & "|" & .Fields(16).Value & "|" & .Fields(17).Value & "|" & .Fields(18).Value & "|" & .Fields(19).Value & "|" & .Fields(20).Value & "|" & .Fields(21).Value & "|" & .Fields(22).Value & "|" & .Fields(23).Value & "|" & .Fields(24).Value & "|" & .Fields(25).Value & "|" & .Fields(26).Value & "|" & .Fields(27).Value & "|" & .Fields(28).Value & "|" & .Fields(29).Value & "|" & .Fields(30).Value & "|" & .Fields(31).Value & "|" & .Fields(32).Value & "|" & .Fields(33).Value & "|" & .Fields(34).Value & "|"
                            
                            .Fields(35).Value = 1
                            .Update
                            .MoveNext
                        Wend
                        
                        Close #1
                    End If
                    
                    .Close
                End With
                
                With CN
                    .Close
                End With
            
                Unload Me
            End If
    End Select
End Sub

Private Sub Command2_Click()
    With Combo1
        If .Text = "" Then
            MsgBox "Por favor seleccione el lugar de origen", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text1(0)
        If .Text = "" Then
            MsgBox "Por introduzca la cantidad de capotes", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If

        If Val(.Text) <= 0 Then
            MsgBox "La cantidad de capotes ingresada no es v�lida, por favor corrija la informaci�n", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
        
        If Val(.Text) <> Round(Val(.Text)) Then
            MsgBox "La cantidad de capotes ingresada debe ser un n�mero entero, por favor corrija la informaci�n", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text1(1)
        If .Text = "" Then
            MsgBox "Por introduzca la cantidad de kilogramos", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    
        If Val(.Text) <= 0 Then
            MsgBox "La cantidad de kilogramos ingresada no es v�lida, por favor corrija la informaci�n", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    Dim CN As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Dim vWip_entity_id As Double
    Dim vOrganization_id As Integer
    Dim vDescription As String
    Dim vPrimary_item_id As Double
    Dim vScheduled_start_date As Date
    Dim vQuantity_completed As String
    Dim vJob_name As String
    Dim vOrg_id As Integer
    Dim vTransfer_status As Integer
    Dim vSubinventory_code As String
    Dim vTransaction_type_id As Integer
    Dim vTransaction_action_id As Integer
    Dim vTransaction_source_type_id As Integer
    Dim vTransaction_uom As String
    Dim vInventory_item_id As Double
    
    Dim Str As String
    Dim ArrStr() As String
    
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
    End With
    
    Str = Combo1.Text
    ArrStr() = Split(Str, "[")
    
    vOrganization_id = Replace(ArrStr(1), "]", "")
    vDescription = "Capote (Kilogramos)"
    vPrimary_item_id = 1113144
    vScheduled_start_date = Date
    vQuantity_completed = Text1(1).Text
    vOrg_id = 490
    vTransfer_status = 0
    vSubinventory_code = "MP"
    vTransaction_source_type_id = 5
    vInventory_item_id = 1107152
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select count (*) + 1 from WIP_DISCRETE_JOBS", CN, adOpenStatic, adLockOptimistic
        .Requery
        
        vWip_entity_id = .Fields(0).Value
        vJob_name = "PNVC-" & .Fields(0).Value
        
        .Close
    End With
     
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from WIP_DISCRETE_JOBS", CN, adOpenStatic, adLockOptimistic
        .AddNew
            .Fields(0) = vWip_entity_id
            .Fields(1) = vOrganization_id
            .Fields(2) = vDescription
            .Fields(3) = vPrimary_item_id
            .Fields(4) = vScheduled_start_date
            .Fields(5) = vQuantity_completed
            .Fields(6) = vJob_name
            .Fields(7) = vOrg_id
            .Fields(8) = vTransfer_status
        .Update
        .Requery
        .Close
    End With
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from MTL_MATERIAL_TRANSACTIONS", CN, adOpenStatic, adLockOptimistic
        
        vTransaction_type_id = 44
        vTransaction_action_id = 31
        vTransaction_uom = "KG"
        
        .AddNew
            .Fields(1) = vPrimary_item_id
            .Fields(2) = vOrganization_id
            .Fields(3) = vSubinventory_code
            .Fields(5) = vTransaction_type_id
            .Fields(6) = vTransaction_action_id
            .Fields(7) = vTransaction_source_type_id
            .Fields(8) = vWip_entity_id
            .Fields(9) = vJob_name
            .Fields(10) = vQuantity_completed
            .Fields(11) = vTransaction_uom
            .Fields(12) = vScheduled_start_date
        .Update
        .Requery
        
        vTransaction_type_id = 35
        vTransaction_action_id = 1
        vTransaction_uom = "PZA"
        vQuantity_completed = Text1(0).Text
        
        .AddNew
            .Fields(1) = vInventory_item_id
            .Fields(2) = vOrganization_id
            .Fields(3) = vSubinventory_code
            .Fields(5) = vTransaction_type_id
            .Fields(6) = vTransaction_action_id
            .Fields(7) = vTransaction_source_type_id
            .Fields(8) = vWip_entity_id
            .Fields(9) = vJob_name
            .Fields(10) = "-" & vQuantity_completed
            .Fields(11) = vTransaction_uom
            .Fields(12) = vScheduled_start_date
        .Update
        .Requery
        .Close
    End With
    
    With CN
        .Close
    End With
    
    MsgBox "Transacci�n guardada con �xito", vbOKOnly, "Terminado"
    
    With Frame1(0)
        .Visible = False
    End With
    
    With Image1
        .Visible = True
    End With
End Sub

Private Sub Command3_Click()
    With Combo2(0)
        If .Text = "" Then
            MsgBox "Por favor seleccione el lugar de origen", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Combo2(1)
        If .Text = "" Then
            MsgBox "Por favor seleccione el producto", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text2(0)
        If .Text = "" Then
            MsgBox "Por introduzca la cantidad de kilogramos", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
        
        If Val(.Text) <= 0 Then
            MsgBox "La cantidad de kilogramos ingresada no es v�lida, por favor corrija la informaci�n", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text2(1)
        If .Text = "" Then
            MsgBox "Por introduzca el n�mero de lote", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    Dim CN As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Dim vWip_entity_id As Double
    Dim vOrganization_id As Integer
    Dim vDescription As String
    Dim vPrimary_item_id As Double
    Dim vScheduled_start_date As Date
    Dim vQuantity_completed As String
    Dim vJob_name As String
    Dim vOrg_id As Integer
    Dim vTransfer_status As Integer
    Dim vSubinventory_code As String
    Dim vTransaction_type_id As Integer
    Dim vTransaction_action_id As Integer
    Dim vTransaction_source_type_id As Integer
    Dim vTransaction_uom As String
    Dim vInventory_item_id As Double
    Dim vLote As String
    
    Dim Str As String
    Dim ArrStr() As String
    
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
    End With
    
    Str = Combo2(0).Text
    ArrStr() = Split(Str, "[")
    
    vOrganization_id = Replace(ArrStr(1), "]", "")
    
    Str = Combo2(1).Text
    ArrStr() = Split(Str, "[")
    
    vDescription = Trim(ArrStr(0))
    vPrimary_item_id = Replace(ArrStr(2), "]", "")
    vScheduled_start_date = Date
    vQuantity_completed = Text2(0).Text
    vLote = Text2(1).Text
    vOrg_id = 490
    vTransfer_status = 0
    vTransaction_source_type_id = 5
    vInventory_item_id = 1113144
    vTransaction_uom = "KG"
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select count (*) + 1 from WIP_DISCRETE_JOBS", CN, adOpenStatic, adLockOptimistic
        .Requery
        
        vWip_entity_id = .Fields(0).Value
        vJob_name = "PNVC-" & .Fields(0).Value
        
        .Close
    End With
     
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from WIP_DISCRETE_JOBS", CN, adOpenStatic, adLockOptimistic
        .AddNew
            .Fields(0) = vWip_entity_id
            .Fields(1) = vOrganization_id
            .Fields(2) = vDescription
            .Fields(3) = vPrimary_item_id
            .Fields(4) = vScheduled_start_date
            .Fields(5) = vQuantity_completed
            .Fields(6) = vJob_name
            .Fields(7) = vOrg_id
            .Fields(8) = vTransfer_status
        .Update
        .Requery
        .Close
    End With
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from MTL_MATERIAL_TRANSACTIONS", CN, adOpenStatic, adLockOptimistic
        
        vTransaction_type_id = 44
        vTransaction_action_id = 31
        vSubinventory_code = "PT"
        
        .AddNew
            .Fields(1) = vPrimary_item_id
            .Fields(2) = vOrganization_id
            .Fields(3) = vSubinventory_code
            .Fields(5) = vTransaction_type_id
            .Fields(6) = vTransaction_action_id
            .Fields(7) = vTransaction_source_type_id
            .Fields(8) = vWip_entity_id
            .Fields(9) = vJob_name
            .Fields(10) = vQuantity_completed
            .Fields(11) = vTransaction_uom
            .Fields(12) = vScheduled_start_date
            .Fields(20) = vLote 'ATTIBUTE1
        .Update
        .Requery
        
        vTransaction_type_id = 35
        vTransaction_action_id = 1
        vSubinventory_code = "MP"
        
        .AddNew
            .Fields(1) = vInventory_item_id
            .Fields(2) = vOrganization_id
            .Fields(3) = vSubinventory_code
            .Fields(5) = vTransaction_type_id
            .Fields(6) = vTransaction_action_id
            .Fields(7) = vTransaction_source_type_id
            .Fields(8) = vWip_entity_id
            .Fields(9) = vJob_name
            .Fields(10) = "-" & vQuantity_completed
            .Fields(11) = vTransaction_uom
            .Fields(12) = vScheduled_start_date
        .Update
        .Requery
        .Close
    End With
    
    With CN
        .Close
    End With
    
    MsgBox "Transacci�n guardada con �xito", vbOKOnly, "Terminado"
    
    With Frame1(1)
        .Visible = False
    End With
    
    With Image1
        .Visible = True
    End With
End Sub

Private Sub Command4_Click()
    With Combo3(0)
        If .Text = "" Then
            MsgBox "Por favor seleccione el lugar de origen", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Combo3(1)
        If .Text = "" Then
            MsgBox "Por favor seleccione el lugar de destino", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Combo3(2)
        If .Text = "" Then
            MsgBox "Por favor seleccione el producto", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text3
        If .Text = "" Then
            MsgBox "Por introduzca la cantidad de kilogramos", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
        
        If Val(.Text) <= 0 Then
            MsgBox "La cantidad de kilogramos ingresada no es v�lida, por favor corrija la informaci�n", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text4
        If .Text = "" Then
            MsgBox "Por introduzca el cliente", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    With Text5
        If .Text = "" Then
            MsgBox "Por introduzca el lote", vbOKOnly, "Advertencia"
            
            .SetFocus
            
            Exit Sub
        End If
    End With
    
    If Combo3(0).Text = Combo3(1).Text Then
        MsgBox "El lugar de origen y el lugar de destino no pueden ser iguales, por favor corrija la informaci�n", vbOKOnly, "Advertencia"
            
        Combo3(0).SetFocus
            
        Exit Sub
    End If
    
    For i = 0 To 1
        With Combo3(i)
            .Enabled = False
        End With
    Next i
    
    With Text4
        .Enabled = False
    End With
    
    
    Dim CN As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    Dim vInventory_item_id As Double
    Dim vOrganization_id As Integer
    Dim vSubinventory_code As String
    Dim vTransaction_type_id As Integer
    Dim vTransaction_quantity As String
    Dim vTransaction_uom As String
    Dim vTransaction_date As Date
    Dim vTransfer_organization_id As Integer
    Dim vTransfer_subinventory As String
    Dim vArticulo As String
    Dim vLote As String
    
    Dim Str As String
    Dim ArrStr() As String
    
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
    End With
    
    Str = Combo3(2).Text
    ArrStr() = Split(Str, "[")
    
    vInventory_item_id = Replace(ArrStr(2), "]", "")
    vArticulo = Replace(ArrStr(0), "]", "")
    
    Str = Combo3(1).Text
    ArrStr() = Split(Str, "[")
    
    vOrganization_id = Replace(ArrStr(1), "]", "")
    vSubinventory_code = "PT"
    vTransaction_type_id = 3
    vTransaction_quantity = Text3.Text
    vTransaction_uom = "KG"
    vTransaction_date = Date
    vLote = Text5.Text
    
    Str = Combo3(0).Text
    ArrStr() = Split(Str, "[")
    
    vTransfer_organization_id = Replace(ArrStr(1), "]", "")
    vTransfer_subinventory = "PT"
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from MTL_MATERIAL_TRANSACTIONS", CN, adOpenStatic, adLockOptimistic
        .AddNew
            .Fields(1) = vInventory_item_id
            .Fields(2) = vOrganization_id
            .Fields(3) = vSubinventory_code
            .Fields(5) = vTransaction_type_id
            .Fields(10) = vTransaction_quantity
            .Fields(11) = vTransaction_uom
            .Fields(12) = vTransaction_date
            .Fields(16) = vTransfer_organization_id
            .Fields(17) = vTransfer_subinventory
            .Fields(20) = vLote
        .Update
        .Requery
        .Close
    End With
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from TICKET", CN, adOpenStatic, adLockOptimistic
        .AddNew
            .Fields(1) = Text4.Text
            .Fields(2) = vArticulo
            .Fields(3) = vTransaction_quantity
        .Update
        .Requery
        .Close
    End With
    
    With CN
        .Close
    End With
    
    MsgBox "Transacci�n guardada con �xito", vbOKOnly, "Terminado"
    
    With Combo3(2)
        .Clear
    End With
            
    With Text3
        .Text = ""
    End With
    
    With Text5
        .Text = ""
    End With
            
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
    End With
            
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from MTL_SYSTEM_ITEMS_B Where ENABLED_FLAG = 'Y'  Order by DESCRIPTION", CN, adOpenStatic, adLockOptimistic
                
        If .RecordCount <> 0 Then
            .MoveFirst
                    
            While Not .EOF
                Combo3(2).AddItem .Fields(2).Value & " [" & .Fields(1).Value & "] [" & .Fields(0).Value & "]"
                        
                .MoveNext
            Wend
        End If
                
        .Close
    End With
            
    With CN
        .Close
    End With
End Sub

Private Sub Command5_Click()
    vbq = MsgBox("�Desea imprimir el ticket y cerrar la pantalla?", vbQuestion + vbYesNo, "Advertencia")
                
    If vbq = vbNo Then
        Exit Sub
    End If

    Dim CN As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    
    With CN
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\BD.mdb; Persist Security Info=False"
    End With
    
    With RS
        .CursorLocation = adUseClient
        .Open "Select * from TICKET where cliente = '" & Text4.Text & "' and impreso = 0", CN, adOpenStatic, adLockOptimistic
        .Requery
    End With
    
    With DataReport1
        Set .DataSource = RS
        
        With .Sections("Section4")
            .Controls("Etiqueta6").Caption = Text4.Text
        End With
        
        With .Sections("Section1")
            .Controls("Texto1").DataField = "ARTICULO"
            .Controls("Texto5").DataField = "CANTIDAD"
        End With
        
        .Show 1
    End With
    
    With RS
        .Close
    End With
    
    With CN
        .Execute "Update TICKET set impreso = 1 where cliente = '" & Text4.Text & "' and impreso = 0"
        .Close
    End With
    
    For i = 0 To 1
        With Combo3(i)
            .Enabled = True
        End With
    Next i
    
    With Text4
        .Enabled = True
    End With
    
    With Frame1(2)
        .Visible = False
    End With
    
    With Image1
        .Visible = True
    End With
End Sub
