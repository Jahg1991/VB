VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas"
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
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14055
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   13815
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   1320
            TabIndex        =   11
            Top             =   960
            Width           =   11775
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   13200
            MaskColor       =   &H009EC0C2&
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
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
            Index           =   3
            Left            =   1320
            MaxLength       =   13
            TabIndex        =   9
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   4440
            MaxLength       =   6
            TabIndex        =   8
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "Agregar"
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
            Index           =   2
            Left            =   120
            MaskColor       =   &H009EC0C2&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "Quitar"
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
            Index           =   3
            Left            =   1440
            MaskColor       =   &H009EC0C2&
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   11760
            TabIndex        =   5
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   1320
            TabIndex        =   4
            Top             =   600
            Width           =   11775
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H009EC0C2&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   13200
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Width           =   495
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1320
            TabIndex        =   12
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
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   10404034
            CalendarTrailingForeColor=   10404034
            Format          =   129302529
            CurrentDate     =   43810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Folio"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   10560
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Articulo"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   13
            Top             =   1320
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009EC0C2&
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   14055
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3015
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   13815
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
            ItemData        =   "Form4.frx":10CA
            Left            =   120
            List            =   "Form4.frx":10CC
            TabIndex        =   23
            Top             =   120
            Width           =   13575
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   11520
            TabIndex        =   22
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H009EC0C2&
            Caption         =   "Guardar"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   9840
            TabIndex        =   21
            Top             =   2640
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

    Dim listTotal As Double

    Select Case Index
        Case 0
            TipoCatalogo = 1
            Form2.Show
            Form2.Text1.SetFocus
        Case 1
            TipoCatalogo = 0
            Form2.Show
            Form2.Text1.SetFocus
        Case 2
            If Text1(2) <> "" And Text1(3) <> "" And Text1(4) <> "" Then
                'JAHG Longitud Cadena
                Dim viid As String
                Dim viarticulo As String
                Dim vicantidad As String
                Dim viprecio As String
                Dim visubtotal As String
                Dim c1 As Integer
                Dim c2 As Integer
                Dim c3 As Integer
                Dim c4 As Integer
                Dim c5 As Integer
                viid = IdItem
                viarticulo = Mid(Text1(2), 1, 70)
                vicantidad = Round(Text1(3), 2)
                viprecio = Round(Text1(4), 2)
                visubtotal = Round(Text1(3) * Text1(4), 2)
                'MsgBox Text1(2) * Text1(3)
                ' 1 - 5
                c1 = 4 - Len(viid)
                For i = 1 To c1
                    viid = " " & viid
                Next i
                ' 6 - 75 (Codigo 6 - 30 , descripcion 31 - 80)
                c2 = 70 - Len(viarticulo)
                For i = 1 To c2
                    viarticulo = viarticulo & " "
                Next i
                ' 76 - 90
                c3 = 15 - Len(vicantidad)
                For i = 1 To c3
                    vicantidad = " " & vicantidad
                Next i
                ' 91 - 105
                c4 = 8 - Len(viprecio)
                For i = 1 To c4
                    viprecio = " " & viprecio
                Next i
                ' 106 - 125
                c5 = 20 - Len(visubtotal)
                For i = 1 To c5
                    visubtotal = " " & visubtotal
                Next i
                List1.AddItem viid & " " & viarticulo & vicantidad & "KG $ " & viprecio & " $" & visubtotal
                Text1(2).Text = ""
                Text1(3).Text = ""
                Text1(4).Text = ""
                Command1(1).SetFocus
                IdItem = ""
                listTotal = 0
                For i = 0 To List1.ListCount - 1
                    List1.ListIndex = i
                    List1.SetFocus
                    'MsgBox Trim(Mid(List1.Text, 106, 20))
                    listTotal = Trim(Mid(List1.Text, 106, 20)) + listTotal
                Next i
                Text1(5) = "$ " & listTotal
            Else
                MsgBox "Llene todos los campos", vbOKOnly, "Advertencia"
            End If
        Case 3
            Dim intX As Integer
            intX = List1.ListIndex
            List1.RemoveItem (intX)
            listTotal = 0
            For i = 0 To List1.ListCount - 1
                List1.ListIndex = i
                List1.SetFocus
                listTotal = Trim(Mid(List1.Text, 106, 20)) + listTotal
            Next i
            Text1(5) = listTotal
    End Select
    
End Sub

Private Sub Command2_Click()

    If IdCliente = "" Then
        MsgBox "Elija un cliente", vbOKOnly, "Advertencia"
    Else
        If List1.ListCount = 0 Then
            MsgBox "Agregue por lo menos un articulo", vbOKOnly, "Advertencia"
        Else
            With RsVentas
                For i = 0 To List1.ListCount - 1
                    
                    List1.ListIndex = i
                    List1.SetFocus
                    .AddNew
                        .Fields("Id") = IdTransaccion
                        .Fields("Folio") = Text1(0)
                        .Fields("Fecha") = DTPicker1.Value
                        .Fields("Cliente") = IdCliente
                        .Fields("Item") = Trim(Mid(List1.Text, 1, 4))
                        .Fields("Cantidad") = Trim(Mid(List1.Text, 76, 15))
                        .Fields("Precio") = Trim(Mid(List1.Text, 96, 8))
                        .Fields("Iva") = 0
                        .Fields("Total") = Trim(Mid(List1.Text, 106, 20))
                        .Fields("Lugar") = "CDMX"
                    .Update
                    .Requery
                Next i
            End With
            'Inicio JAHG TXT
            If PrERP = 1 Then
                Dim myFile As String, Nombre_csv As String, Ruta_csv As String, RutaCompleta_csv As String
                tmp_val = ""
                RsVentasTxt.Requery
                RsVentasTxt.Filter = "Folio like  '" & Text1(0) & "'"
                If Not RsVentasTxt.EOF Then
                    RsVentasTxt.MoveLast
                    rcount = RsVentasTxt.RecordCount
                    RsVentasTxt.MoveFirst
                    Nombre_csv = Text1(0) & ".txt"
                    Ruta_csv = App.Path
                    RutaCompleta_csv = Ruta_csv & "\ERP\"
                    myFile = RutaCompleta_csv & Nombre_csv
                    Open myFile For Output As #1
                    While Not RsVentasTxt.EOF
                        For i = 0 To RsVentasTxt.Fields.Count - 1
                        tmp_val = tmp_val & RsVentasTxt.Fields(RsVentasTxt.Fields(i).Name) & "|"
                        Next i
                        tmp_val = Mid(tmp_val, 1, Len(tmp_val) - 1)
                        tmp_val = tmp_val & vbCrLf
                        RsVentasTxt.MoveNext
                        DoEvents
                    Wend
                    Print #1, tmp_val
                End If
                Close #1
            End If
            'Fin JAHG TXT
            'Inicio Ticket
            '26 Caracteres por linea
            If PrTicket = 1 Then
                Dim sMensaje As String
                Dim linea As Integer
                linea = 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = Mid(PrEmpresa, 1, 26) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & Mid(PrRfc, 1, 26) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & Mid(PrCiudad & ", " & PrEstado, 1, 26) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "C.P. " & Mid(PrCodigoPostal, 1, 5) & "Tel. " & Mid(PrTelefono, 1, 10) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & Mid(PrCorreo, 1, 26) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "--------------------------" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "Cliente " & Mid(Text1(1), 1, 18) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "Fecha " & DTPicker1.Value & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "Folio " & Text1(0) & vbNewLine
                For i = 0 To List1.ListCount - 1
                    List1.ListIndex = i
                    List1.SetFocus
                    linea = linea + 1
                    sMensaje = sMensaje & "==========================" & vbNewLine
                    linea = linea + 1
                    sMensaje = sMensaje & "Art. " & Trim(Mid(List1.Text, 31, 21)) & vbNewLine
                    linea = linea + 1
                    sMensaje = sMensaje & "Cant. " & Trim(Mid(List1.Text, 76, 15)) & "kg" & vbNewLine
                    linea = linea + 1
                    sMensaje = sMensaje & "Precio u. " & Trim(Mid(List1.Text, 96, 8)) & vbNewLine
                    linea = linea + 1
                    sMensaje = sMensaje & "Sub. " & Trim(Mid(List1.Text, 106, 20)) & vbNewLine
                Next i
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "==========================" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "A pagar " & Text1(5) & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "==========================" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
    
                linea = linea + 1
                sMensaje = sMensaje & "--------------------------" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "  Gracias por su compra" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "**************************" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "--------------------------" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & "          Firma" & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                linea = linea + 1
                sMensaje = sMensaje & vbNewLine
                Unload Form7
                With Form7
                    .Height = 1000 + (linea * 300)
                    .FontSize = 10
                    .FontName = "CONSOLAS"
                    .CurrentY = 300
                    Form7.Print sMensaje
                    .Show 1
                End With
                'Fin Ticket
            End If
            DTPicker1.Value = Date
            With RsIdVenta
                .Requery
                .MoveFirst
                If IsNull(RsIdVenta!IdTransaccion) = False Then
                    IdTransaccion = RsIdVenta!IdTransaccion
                Else
                    IdTransaccion = 1
                End If
            End With
            Text1(0).Text = "CDMX-" & IdTransaccion
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
            Text1(5).Text = ""
            List1.Clear
            IdCliente = ""
            IdItem = ""
        End If
    End If

End Sub

Private Sub Form_Load()

    If PrPrecio = 1 Then
        Text1(4).Enabled = True
    Else
        Text1(4).Enabled = False
    End If
    DTPicker1.Value = Date
    With RsIdVenta
        .Requery
        .MoveFirst
        If IsNull(RsIdVenta!IdTransaccion) = False Then
            IdTransaccion = RsIdVenta!IdTransaccion
        Else
            IdTransaccion = 1
        End If
    End With
    Text1(0).Text = "DSJ-" & IdTransaccion
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    IdCliente = ""
    IdItem = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    TipoTransaccion = ""
    IdCliente = ""
    IdItem = ""
    Form1.Enabled = True
    
End Sub
