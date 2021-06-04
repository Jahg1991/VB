VERSION 5.00
Begin VB.Form Form2 
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   13575
         Begin VB.CommandButton Command2 
            BackColor       =   &H009EC0C2&
            Caption         =   "Añadir"
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
            Left            =   12480
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   975
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5070
            ItemData        =   "Form2.frx":0000
            Left            =   120
            List            =   "Form2.frx":0002
            OLEDragMode     =   1  'Automatic
            TabIndex        =   4
            Top             =   600
            Width           =   13335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "Listo"
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
            Left            =   12480
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   12255
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If TipoCatalogo = 0 Then
        Text1 = Trim(Mid(List1.Text, 1, 25))
        RsItems.MoveFirst
        IdItem = RsItems!ID
        If TipoTransaccion = 1 Then
            'JAHG Longitud Cadena
            Dim micadena As String
            Dim micadena1 As String
            Dim caracteres As Integer
            micadena = RsItems!Codigo
            micadena1 = RsItems!Descripcion
            caracteres = 25 - Len(micadena)
            For i = 1 To caracteres
                micadena = micadena & " "
            Next i
            Form4.Text1(2) = micadena & micadena1
            Form4.Text1(4) = RsItems!Precio
            Form4.Text1(3).SetFocus
        End If
    End If
    If TipoCatalogo = 1 Then
        Text1 = List1.Text
        RsClientes.MoveFirst
        IdCliente = RsClientes!ID
        If TipoTransaccion = 0 Then
            Form3.Text1(1) = RsClientes!Nombre
        End If
        If TipoTransaccion = 1 Then
            Form4.Text1(1) = RsClientes!Nombre
            Form4.Command1(1).SetFocus
        End If
    End If
    Unload Me

End Sub

Private Sub Command2_Click()

    Form8.Show

End Sub

Private Sub Form_Load()

    List1.Clear
    If TipoConsulta = 0 Then
        Command1.Visible = False
        'Añadir
        If PrArticulo = 1 Then
            Command2.Visible = True
        Else
            Command2.Visible = False
        End If
    End If
    If TipoConsulta = 1 Then
        If PrCliente = 1 Then
            Command1.Visible = True
        Else
            Command2.Visible = False
        End If
    End If
    If TipoCatalogo = 0 Then
        Form2.Icon = LoadPicture(App.Path & "\Images\Articulos.ico")
        Form2.Caption = "Articulos"
        IdItem = ""
        RsItems.Requery
        RsItems.Filter = ""
        Do Until RsItems.EOF
            'JAHG Longitud Cadena
            Dim micadena As String
            Dim micadena1 As String
            Dim caracteres As Integer
            micadena = RsItems!Codigo
            micadena1 = RsItems!Descripcion
            caracteres = 25 - Len(micadena)
            For i = 1 To caracteres
                micadena = micadena & " "
            Next i
            List1.AddItem micadena & micadena1
            RsItems.MoveNext
        Loop
    End If
    If TipoCatalogo = 1 Then
    Form2.Icon = LoadPicture(App.Path & "\Images\Clientes.ico")
        Form2.Caption = "Clientes"
        RsClientes.Requery
        RsClientes.Filter = ""
        Do Until RsClientes.EOF
            List1.AddItem RsClientes!Nombre
            RsClientes.MoveNext
        Loop
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If TipoCatalogo = 0 Then
        RsItems.Filter = ""
    End If
    If TipoCatalogo = 1 Then
        RsClientes.Filter = ""
    End If
    Form1.Enabled = True

End Sub

Private Sub List1_DblClick()
    
    If TipoConsulta = 1 Then
        If TipoCatalogo = 0 Then
            Text1 = Trim(Mid(List1.Text, 1, 25))
            RsItems.MoveFirst
            IdItem = RsItems!ID
            If TipoTransaccion = 1 Then
                'JAHG Longitud Cadena
                Dim micadena As String
                Dim micadena1 As String
                Dim caracteres As Integer
                micadena = RsItems!Codigo
                micadena1 = RsItems!Descripcion
                caracteres = 25 - Len(micadena)
                For i = 1 To caracteres
                    micadena = micadena & " "
                Next i
                Form4.Text1(2) = micadena & micadena1
                Form4.Text1(4) = RsItems!Precio
                Form4.Text1(3).SetFocus
            End If
        End If
        If TipoCatalogo = 1 Then
            Text1 = List1.Text
            RsClientes.MoveFirst
            IdCliente = RsClientes!ID
            If TipoTransaccion = 0 Then
                Form3.Text1(1) = RsClientes!Nombre
            End If
            If TipoTransaccion = 1 Then
                Form4.Text1(1) = RsClientes!Nombre
                Form4.Command1(1).SetFocus
            End If
        End If
        Unload Me
    End If

End Sub

Private Sub Text1_Change()

    If TipoCatalogo = 0 Then
        Dim micadena As String
        Dim micadena1 As String
        Dim caracteres As Integer
        If Text1 <> "" Then
            RsItems.Filter = "Descripcion like  '*" & Text1 & "*'" & " OR Codigo like  '*" & Text1 & "*'"
            List1.Clear
            Do Until RsItems.EOF
                'JAHG Longitud Cadena
                micadena = RsItems!Codigo
                micadena1 = RsItems!Descripcion
                caracteres = 25 - Len(micadena)
                For i = 1 To caracteres
                    micadena = micadena & " "
                Next i
                List1.AddItem micadena & micadena1
                RsItems.MoveNext
            Loop
        Else
            RsItems.Filter = ""
            List1.Clear
            Do Until RsItems.EOF
                'JAHG Longitud Cadena
                micadena = RsItems!Codigo
                micadena1 = RsItems!Descripcion
                caracteres = 25 - Len(micadena)
                For i = 1 To caracteres
                    micadena = micadena & " "
                Next i
                List1.AddItem micadena & micadena1
                RsItems.MoveNext
            Loop
        End If
    End If
    If TipoCatalogo = 1 Then
        If Text1 <> "" Then
            RsClientes.Filter = "Nombre like  '*" & Text1 & "*'"
            List1.Clear
            Do Until RsClientes.EOF
                List1.AddItem RsClientes!Nombre
                RsClientes.MoveNext
            Loop
        Else
            RsClientes.Filter = ""
            List1.Clear
            Do Until RsClientes.EOF
                List1.AddItem RsClientes!Nombre
                RsClientes.MoveNext
            Loop
        End If
    End If

End Sub
