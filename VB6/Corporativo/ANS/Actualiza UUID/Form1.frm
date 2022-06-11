VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexion EDIX - EBS (Actualiza UUID)"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   10770
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
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
   ScaleHeight     =   5445
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   1905
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3360
      Width           =   7935
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2520
      TabIndex        =   10
      Text            =   "Desconectado"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2520
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   240
      Top             =   4560
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "LOG"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ESTATUS"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BASE DE DATOS"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVIDOR BD"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Menu Guardar 
      Caption         =   "Guardar"
   End
   Begin VB.Menu Conectar 
      Caption         =   "Conectar"
   End
   Begin VB.Menu Desconectar 
      Caption         =   "Desconectar"
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CNSS As New ADODB.Connection
Dim RSSS As New ADODB.Recordset

Dim CNMA As New ADODB.Connection
Dim RSMA As New ADODB.Recordset

Private Sub Conectar_Click()
    On Error Resume Next

    Text5.Text = "Conectado"

    Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " LEYENDO INFORMACIÓN DE BD..."

    If Len(Text6) <= 32500 Then
        Text6.SelStart = Len(Text6)
    Else
        Text6.Text = ""
        Text6.SelStart = Len(Text6)
    End If

    Timer1.Enabled = True
End Sub

Private Sub Desconectar_Click()
    On Error Resume Next

    Text5.Text = "Desonectado"

    Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " DESCONEXIÓN A LA BD CORRECTA"

    If Len(Text6) <= 32500 Then
        Text6.SelStart = Len(Text6)
    Else
        Text6.Text = ""
        Text6.SelStart = Len(Text6)
    End If

    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo ERR
    
    Text5.Text = "Desconectado"
    
    Desconectar.Enabled = False

    Text6.Text = Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " INICIANDO PROGRAMA"
    Text6.SelStart = Len(Text6)

    With RSMA
        If .State = 1 Then .Close
    End With

    With CNMA
        If .State = 1 Then .Close
    End With

    Set RSMA = Nothing
    Set CNMA = Nothing

    With CNMA
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\CNN.mdb; Persist Security Info=False"
    End With

    Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " OBTENIENDO DATOS DE CONEXIÓN"

    If Len(Text6) <= 32500 Then
        Text6.SelStart = Len(Text6)
    Else
        Text6.Text = ""
        Text6.SelStart = Len(Text6)
    End If

    With RSMA
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "SELECT * FROM CNN;", CNMA, adOpenStatic, adLockOptimistic
        .Requery

        If .RecordCount <> 0 Then
            .MoveFirst

            With Text1
                Set .DataSource = RSMA
                .DataField = "SERVIDOR"
            End With

            With Text2
                Set .DataSource = RSMA
                .DataField = "BD"
            End With

            With Text3
                Set .DataSource = RSMA
                .DataField = "USUARIO"
            End With

            With Text4
                Set .DataSource = RSMA
                .DataField = "PASS"
            End With
        End If
    End With
    Exit Sub

ERR:
    If ERR.Number = -2147467259 Then
        Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " ERROR AL OBTENER DATOS DE CONEXION"

        If Len(Text6) <= 32500 Then
            Text6.SelStart = Len(Text6)
        Else
            Text6.Text = ""
            Text6.SelStart = Len(Text6)
        End If

        With RSMA
            If .State = 1 Then .Close
        End With

        With CNMA
            If .State = 1 Then .Close
        End With

        Set RSMA = Nothing
        Set CNMA = Nothing
    End If

    If ERR.Number = -2147217865 Then
        Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " NO SE PUDIERON LEER DATOS DE CONEXIÓN"

        If Len(Text6) <= 32500 Then
            Text6.SelStart = Len(Text6)
        Else
            Text6.Text = ""
            Text6.SelStart = Len(Text6)
        End If

        With RSMA
            If .State = 1 Then .Close
        End With

        With CNMA
            If .State = 1 Then .Close
        End With

        Set RSMA = Nothing
        Set CNMA = Nothing
    End If

    ERR.Clear
End Sub

Private Sub Guardar_Click()
    On Error Resume Next

    With RSMA
        .Update
        .Requery
    End With

    Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " CAMBIOS GUARDADOS CON ÉXITO"

    If Len(Text6) <= 32500 Then
        Text6.SelStart = Len(Text6)
    Else
        Text6.Text = ""
        Text6.SelStart = Len(Text6)
    End If
End Sub

Private Sub Salir_Click()
    On Error Resume Next

    Timer1.Enabled = False

    With RSMA
        If .State = 1 Then .Close
    End With

    With CNMA
        If .State = 1 Then .Close
    End With

    With RSSS
        If .State = 1 Then .Close
    End With

    With CNSS
        If .State = 1 Then .Close
    End With

    Set RSMA = Nothing
    Set CNMA = Nothing

    Set RSSS = Nothing
    Set CNSS = Nothing

    Unload Me
End Sub

Private Sub Text5_Change()
    On Error Resume Next

    If Text5.Text = "Conectado" Then
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Text4.Enabled = False
        Guardar.Enabled = False
        Conectar.Enabled = False
        Desconectar.Enabled = True
    Else
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Text4.Enabled = True
        Guardar.Enabled = True
        Conectar.Enabled = True
        Desconectar.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ERR

    With RSSS
        If .State = 1 Then .Close
    End With

    With CNSS
        If .State = 1 Then .Close
    End With

    Set RSSS = Nothing
    Set CNSS = Nothing

    With CNSS
        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
        If .State = 0 Then .Open "Driver={ODBC Driver 17 for SQL Server};Server=" & Text1 & ";Database=" & Text2 & ";UID=" & Text3 & ";PWD=" & Text4 & ";"
    End With

    With RSSS
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Select a.procesoid,a.cfdi_uuid,REPLACE(convert(varchar,a.cfdi_fechatimbrado,6),' ','-') cfdi_fechatimbrado,a.folio,a.serie_comprobante,rfc.rfc,b.tipocfd From tbl_fak_cfd_movimiento a, tbl_fak_cfd b, tbl_fak_entidad rfc Where a.procesoid = b.procesoid AND b.entidadid_emisor = rfc.entidadid AND a.cfdi_uuid is NOT NULL and NOT EXISTS (Select 1 From Tbl_Ebs_Factura_UUID c Where a.procesoid = c.procesoid);", CNSS, adOpenStatic, adLockOptimistic

        If .RecordCount <> 0 Then
            Open "C:\EkoCFD-Buzones\Emision\UUID\UUID_" & Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".txt" For Output As #1

            .MoveFirst

            While Not .EOF
                CNSS.Execute ("Insert into Tbl_Ebs_Factura_UUID (procesoid,cfdi_uuid,cfdi_fechatimbrado,folio,serie_comprobante,rfc,tipocfd,transfer_status) Values (" & .Fields(0).Value & ",'" & .Fields(1).Value & "','" & .Fields(2).Value & "'," & .Fields(3).Value & ",'" & .Fields(4).Value & "','" & .Fields(5).Value & "','" & .Fields(6).Value & "',1);")

                Print #1, .Fields(0).Value & "|" & .Fields(1).Value & "|" & .Fields(2).Value & "|" & .Fields(3).Value & "|" & .Fields(4).Value & "|" & .Fields(5).Value & "|" & .Fields(6).Value & "|"

                .MoveNext
            Wend

            Close #1

            Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " REGISTROS PROCESADOS: " & .RecordCount

            If Len(Text6) <= 32500 Then
                Text6.SelStart = Len(Text6)
            Else
                Text6.Text = ""
                Text6.SelStart = Len(Text6)
            End If
        End If

        Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " SIN REGISTROS PARA PROCESAR"

        If Len(Text6) <= 32500 Then
            Text6.SelStart = Len(Text6)
        Else
            Text6.Text = ""
            Text6.SelStart = Len(Text6)
        End If
    End With

    With RSSS
        If .State = 1 Then .Close
    End With

    With CNSS
        If .State = 1 Then .Close
    End With

    Set RSSS = Nothing
    Set CNSS = Nothing

    Exit Sub

ERR:

    If ERR.Number = -2147467259 Then
        Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " ERROR DE CONEXIÓN A LA BASE DE DATOS"

        If Len(Text6) <= 32500 Then
            Text6.SelStart = Len(Text6)
        Else
            Text6.Text = ""
            Text6.SelStart = Len(Text6)
        End If

        With RSSS
            If .State = 1 Then .Close
        End With

        With CNSS
            If .State = 1 Then .Close
        End With

        Set RSSS = Nothing
        Set CNSS = Nothing

        Timer1.Enabled = False

        Text5.Text = "Desconectado"
    End If

    If ERR.Number = -2147217865 Then
        Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " NO SE ENCONTRARON LAS TABLAS DE DATOS"
        Text6.Text = Text6.Text & vbCrLf & Format(Date, "YYYY/MM/DD") & Format(Time, "HH:MM:SS") & " DESCONECTANDO DE LA BASE DE DATOS"

        If Len(Text6) <= 32500 Then
            Text6.SelStart = Len(Text6)
        Else
            Text6.Text = ""
            Text6.SelStart = Len(Text6)
        End If

        With RSSS
            If .State = 1 Then .Close
        End With

        With CNSS
            If .State = 1 Then .Close
        End With

        Set RSSS = Nothing
        Set CNSS = Nothing

        Timer1.Enabled = False

        Text5.Text = "Desconectado"
    End If

    ERR.Clear
End Sub
