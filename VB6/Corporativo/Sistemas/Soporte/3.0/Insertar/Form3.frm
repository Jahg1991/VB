VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   0
      Left            =   8520
      TabIndex        =   14
      Top             =   2160
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   1
      Left            =   8520
      TabIndex        =   13
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   2
      Left            =   8520
      TabIndex        =   12
      Top             =   3360
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   3
      Left            =   8520
      TabIndex        =   11
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   4
      Left            =   8520
      TabIndex        =   10
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   5
      Left            =   8520
      TabIndex        =   9
      Top             =   5160
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   6
      Left            =   8520
      TabIndex        =   8
      Top             =   5760
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   405
      Left            =   10440
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   8
      Left            =   240
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   405
      Left            =   9600
      TabIndex        =   5
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nuevo"
      Height          =   405
      Left            =   12240
      TabIndex        =   4
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   405
      Left            =   10920
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   7
      Left            =   840
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
      ForeColor       =   4210752
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Àrea"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   21
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Puesto"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   20
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correo interno"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   18
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correo externo"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   4
      Left            =   6480
      TabIndex        =   17
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nùmero de celular"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   5
      Left            =   6480
      TabIndex        =   16
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Extensiòn"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   6
      Left            =   6480
      TabIndex        =   15
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   11760
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6975
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If RSAGENDA.BOF = True And RSAGENDA.EOF And RSAGENDA2.BOF = True And RSAGENDA2.EOF = True Then
        
    Else

        CommonDialog1.ShowOpen
        CommonDialog1.Filter = "Imagenes|*.jpg"
        
        'Copiamos la foto
        If CommonDialog1.FileName = "" Then
        
        Else
            FileCopy CommonDialog1.FileName, "C:\JAHG Software\JAHG Soporte\Empleados\Imagenes\" + Text2(8).Text + ".jpg"
            Text2(7).Text = "C:\JAHG Software\JAHG Soporte\Empleados\Imagenes\" + Text2(8).Text + ".jpg"
            
            Image1.Picture = LoadPicture(Text2(7).Text)
        
        End If
        
    End If
    
End Sub

Private Sub Command2_Click()

    If RSAGENDA.BOF = True And RSAGENDA.EOF And RSAGENDA2.BOF = True And RSAGENDA2.EOF = True Then
    
    Else

        RSAGENDA.Update
        
        msg = MsgBox("Cambios guardados correctamente", vbOKOnly, "Listo!")
    
    End If

End Sub

Private Sub Command3_Click()
    
    Form4.Show
    
    Set Form4.Text2(8).DataSource = RSID
    Form4.Text2(8).DataField = "ID"
    
End Sub

Private Sub Command4_Click()

    If RSAGENDA.BOF = True And RSAGENDA.EOF And RSAGENDA2.BOF = True And RSAGENDA2.EOF = True Then
    
    Else

        RSAGENDA.Delete
        
        msg = MsgBox("Eliminado correctamente", vbOKOnly, "Listo!")
        
        RSAGENDA.Requery
        
        RSAGENDA2.Requery
        
        Text1.Text = ""
        
        Image1.Picture = LoadPicture(Text2(7).Text)
        
    End If

End Sub

Private Sub DataGrid1_Click()

    Text1.Text = DataGrid1.Columns(0).Text

End Sub

Private Sub Text1_Change()

    DataGrid1.Columns(0).Width = 4700
    
    'RSAGENDA.Requery
    
    If Text1 = "" Then
    
        RSAGENDA.Filter = ""
        RSAGENDA2.Filter = ""
        
        DataGrid1.Columns(0).Width = 4700
        
        Image1.Picture = LoadPicture(Text2(7).Text)
        
        If RSAGENDA.BOF = True And RSAGENDA.EOF And RSAGENDA2.BOF = True And RSAGENDA2.EOF = True Then
        
        Else
        
            RSAGENDA.MoveFirst
            RSAGENDA2.MoveFirst
            
                
            Set Text2(0).DataSource = RSAGENDA
            Text2(0).DataField = "AREA"
                    
            Set Text2(1).DataSource = RSAGENDA
            Text2(1).DataField = "PUESTO"
                    
            Set Text2(2).DataSource = RSAGENDA
            Text2(2).DataField = "NOMBRE"
                    
            Set Text2(3).DataSource = RSAGENDA
            Text2(3).DataField = "CORREO_INTERNO"
                    
            Set Text2(4).DataSource = RSAGENDA
            Text2(4).DataField = "CORREO_EXTERNO"
                    
            Set Text2(5).DataSource = RSAGENDA
            Text2(5).DataField = "CELULAR"
                    
            Set Text2(6).DataSource = RSAGENDA
            Text2(6).DataField = "EXTENSION"
                   
            Set Text2(7).DataSource = RSAGENDA
            Text2(7).DataField = "IMAGEN"
                    
            Set DataGrid1.DataSource = RSAGENDA2
            
            DataGrid1.Columns(0).Width = 4700
            
            Image1.Picture = LoadPicture(Text2(7).Text)
            
        End If
    
    Else
    
        RSAGENDA.Filter = "NOMBRE like '*" & Text1 & "*'"
        RSAGENDA2.Filter = "NOMBRE like '*" & Text1 & "*'"
            
        Set Text2(0).DataSource = RSAGENDA
        Text2(0).DataField = "AREA"
                
        Set Text2(1).DataSource = RSAGENDA
        Text2(1).DataField = "PUESTO"
                
        Set Text2(2).DataSource = RSAGENDA
        Text2(2).DataField = "NOMBRE"
                
        Set Text2(3).DataSource = RSAGENDA
        Text2(3).DataField = "CORREO_INTERNO"
                
        Set Text2(4).DataSource = RSAGENDA
        Text2(4).DataField = "CORREO_EXTERNO"
                
        Set Text2(5).DataSource = RSAGENDA
        Text2(5).DataField = "CELULAR"
                
        Set Text2(6).DataSource = RSAGENDA
        Text2(6).DataField = "EXTENSION"
               
        Set Text2(7).DataSource = RSAGENDA
        Text2(7).DataField = "IMAGEN"
                
        Set DataGrid1.DataSource = RSAGENDA2
        
        DataGrid1.Columns(0).Width = 4700
        
        Image1.Picture = LoadPicture(Text2(7).Text)
        
    End If
    
End Sub
