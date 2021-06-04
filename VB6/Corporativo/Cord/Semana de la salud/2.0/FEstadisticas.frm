VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FEstadisticas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadísticas"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   7185
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
   Icon            =   "FEstadisticas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin TabDlg.SSTab SSTab1 
         Height          =   7815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   13785
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Asistentes"
         TabPicture(0)   =   "FEstadisticas.frx":324A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Género"
         TabPicture(1)   =   "FEstadisticas.frx":3266
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3(2)"
         Tab(1).Control(1)=   "Frame2(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Edad"
         TabPicture(2)   =   "FEstadisticas.frx":3282
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3(3)"
         Tab(2).Control(1)=   "Frame2(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Tipo"
         TabPicture(3)   =   "FEstadisticas.frx":329E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame3(0)"
         Tab(3).Control(1)=   "Frame2(3)"
         Tab(3).ControlCount=   2
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Tipo"
            Height          =   7095
            Index           =   3
            Left            =   -73080
            TabIndex        =   162
            Top             =   480
            Width           =   4695
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   21
               Left            =   2400
               TabIndex        =   242
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   20
               Left            =   1680
               TabIndex        =   241
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   240
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   1680
               TabIndex        =   239
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   21
               Left            =   840
               TabIndex        =   238
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   20
               Left            =   120
               TabIndex        =   237
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   19
               Left            =   840
               TabIndex        =   236
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   120
               TabIndex        =   235
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Invitados"
               Height          =   375
               Index           =   16
               Left            =   1680
               TabIndex        =   204
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Trabajadores"
               Height          =   375
               Index           =   18
               Left            =   120
               TabIndex        =   203
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   202
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   20
               Left            =   1680
               TabIndex        =   201
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   21
               Left            =   120
               TabIndex        =   200
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   22
               Left            =   840
               TabIndex        =   199
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   198
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   197
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   196
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   3
               Left            =   840
               TabIndex        =   195
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   194
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   5
               Left            =   840
               TabIndex        =   193
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   192
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   7
               Left            =   840
               TabIndex        =   191
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   190
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   9
               Left            =   840
               TabIndex        =   189
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   188
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   11
               Left            =   840
               TabIndex        =   187
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   186
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   13
               Left            =   840
               TabIndex        =   185
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   184
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   15
               Left            =   840
               TabIndex        =   183
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   120
               TabIndex        =   182
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe8 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   17
               Left            =   840
               TabIndex        =   181
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   1680
               TabIndex        =   180
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   1
               Left            =   2400
               TabIndex        =   179
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   1680
               TabIndex        =   178
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   3
               Left            =   2400
               TabIndex        =   177
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   1680
               TabIndex        =   176
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   5
               Left            =   2400
               TabIndex        =   175
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   1680
               TabIndex        =   174
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   7
               Left            =   2400
               TabIndex        =   173
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   1680
               TabIndex        =   172
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   9
               Left            =   2400
               TabIndex        =   171
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   1680
               TabIndex        =   170
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   11
               Left            =   2400
               TabIndex        =   169
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   12
               Left            =   1680
               TabIndex        =   168
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   13
               Left            =   2400
               TabIndex        =   167
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   1680
               TabIndex        =   166
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   15
               Left            =   2400
               TabIndex        =   165
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   1680
               TabIndex        =   164
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   17
               Left            =   2400
               TabIndex        =   163
               Top             =   5400
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Módulos"
            Height          =   7215
            Index           =   0
            Left            =   -74880
            TabIndex        =   152
            Top             =   360
            Width           =   1815
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cardiología"
               Height          =   375
               Index           =   64
               Left            =   0
               TabIndex        =   244
               Top             =   6480
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tuberculosis"
               Height          =   375
               Index           =   63
               Left            =   0
               TabIndex        =   243
               Top             =   6000
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Somatometría"
               Height          =   375
               Index           =   23
               Left            =   0
               TabIndex        =   161
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laboratorio"
               Height          =   375
               Index           =   24
               Left            =   0
               TabIndex        =   160
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dental"
               Height          =   375
               Index           =   25
               Left            =   0
               TabIndex        =   159
               Top             =   2640
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nutrición"
               Height          =   375
               Index           =   26
               Left            =   0
               TabIndex        =   158
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Optometría"
               Height          =   375
               Index           =   28
               Left            =   0
               TabIndex        =   157
               Top             =   5040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Audiometría"
               Height          =   375
               Index           =   29
               Left            =   0
               TabIndex        =   156
               Top             =   5520
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DOCMA"
               Height          =   375
               Index           =   30
               Left            =   1080
               TabIndex        =   155
               Top             =   3600
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DOCCU"
               Height          =   375
               Index           =   31
               Left            =   1080
               TabIndex        =   154
               Top             =   4080
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mastografía"
               Height          =   375
               Index           =   32
               Left            =   360
               TabIndex        =   153
               Top             =   4560
               Width           =   1455
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Edad"
            Height          =   6855
            Index           =   2
            Left            =   -73080
            TabIndex        =   88
            Top             =   480
            Width           =   4695
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   21
               Left            =   3960
               TabIndex        =   234
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   20
               Left            =   3240
               TabIndex        =   233
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   19
               Left            =   3960
               TabIndex        =   232
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   3240
               TabIndex        =   231
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   21
               Left            =   2400
               TabIndex        =   230
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   20
               Left            =   1680
               TabIndex        =   229
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   228
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   1680
               TabIndex        =   227
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   21
               Left            =   840
               TabIndex        =   226
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   20
               Left            =   120
               TabIndex        =   225
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   19
               Left            =   840
               TabIndex        =   224
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   120
               TabIndex        =   223
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "31-60 años"
               Height          =   375
               Index           =   8
               Left            =   1680
               TabIndex        =   151
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0-30 años"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   150
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   10
               Left            =   2400
               TabIndex        =   149
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   11
               Left            =   1680
               TabIndex        =   148
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   147
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   13
               Left            =   840
               TabIndex        =   146
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "61 años en adelante"
               Height          =   615
               Index           =   14
               Left            =   3240
               TabIndex        =   145
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   15
               Left            =   3240
               TabIndex        =   144
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   17
               Left            =   3960
               TabIndex        =   143
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   142
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   141
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   140
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   3
               Left            =   840
               TabIndex        =   139
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   138
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   5
               Left            =   840
               TabIndex        =   137
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   136
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   7
               Left            =   840
               TabIndex        =   135
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   134
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   9
               Left            =   840
               TabIndex        =   133
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   132
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   11
               Left            =   840
               TabIndex        =   131
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   130
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   13
               Left            =   840
               TabIndex        =   129
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   128
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   15
               Left            =   840
               TabIndex        =   127
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   120
               TabIndex        =   126
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   17
               Left            =   840
               TabIndex        =   125
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   1680
               TabIndex        =   124
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   1
               Left            =   2400
               TabIndex        =   123
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   1680
               TabIndex        =   122
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   3
               Left            =   2400
               TabIndex        =   121
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   1680
               TabIndex        =   120
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   5
               Left            =   2400
               TabIndex        =   119
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   1680
               TabIndex        =   118
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   7
               Left            =   2400
               TabIndex        =   117
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   1680
               TabIndex        =   116
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   9
               Left            =   2400
               TabIndex        =   115
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   1680
               TabIndex        =   114
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   11
               Left            =   2400
               TabIndex        =   113
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   12
               Left            =   1680
               TabIndex        =   112
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   13
               Left            =   2400
               TabIndex        =   111
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   1680
               TabIndex        =   110
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   15
               Left            =   2400
               TabIndex        =   109
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   1680
               TabIndex        =   108
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   17
               Left            =   2400
               TabIndex        =   107
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   3240
               TabIndex        =   106
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   1
               Left            =   3960
               TabIndex        =   105
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   3240
               TabIndex        =   104
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   3
               Left            =   3960
               TabIndex        =   103
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   3240
               TabIndex        =   102
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   5
               Left            =   3960
               TabIndex        =   101
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   3240
               TabIndex        =   100
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   7
               Left            =   3960
               TabIndex        =   99
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   3240
               TabIndex        =   98
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   9
               Left            =   3960
               TabIndex        =   97
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   3240
               TabIndex        =   96
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   11
               Left            =   3960
               TabIndex        =   95
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   12
               Left            =   3240
               TabIndex        =   94
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   13
               Left            =   3960
               TabIndex        =   93
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   3240
               TabIndex        =   92
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   15
               Left            =   3960
               TabIndex        =   91
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   3240
               TabIndex        =   90
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Labe7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   17
               Left            =   3960
               TabIndex        =   89
               Top             =   5400
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Módulos"
            Height          =   6975
            Index           =   3
            Left            =   -74880
            TabIndex        =   78
            Top             =   360
            Width           =   1815
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cardiología"
               Height          =   375
               Index           =   58
               Left            =   0
               TabIndex        =   222
               Top             =   6480
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tuberculosis"
               Height          =   375
               Index           =   48
               Left            =   0
               TabIndex        =   221
               Top             =   6000
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Somatometría"
               Height          =   375
               Index           =   62
               Left            =   0
               TabIndex        =   87
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laboratorio"
               Height          =   375
               Index           =   61
               Left            =   0
               TabIndex        =   86
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dental"
               Height          =   375
               Index           =   60
               Left            =   0
               TabIndex        =   85
               Top             =   2640
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nutrición"
               Height          =   375
               Index           =   59
               Left            =   0
               TabIndex        =   84
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Optometría"
               Height          =   375
               Index           =   57
               Left            =   0
               TabIndex        =   83
               Top             =   5040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Audiometría"
               Height          =   375
               Index           =   56
               Left            =   0
               TabIndex        =   82
               Top             =   5520
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DOCMA"
               Height          =   375
               Index           =   55
               Left            =   1080
               TabIndex        =   81
               Top             =   3600
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DOCCU"
               Height          =   375
               Index           =   54
               Left            =   1080
               TabIndex        =   80
               Top             =   4080
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mastografía"
               Height          =   375
               Index           =   53
               Left            =   360
               TabIndex        =   79
               Top             =   4560
               Width           =   1455
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Género"
            Height          =   5295
            Index           =   1
            Left            =   -73080
            TabIndex        =   43
            Top             =   480
            Width           =   4695
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   220
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   18
               Left            =   1680
               TabIndex        =   219
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   17
               Left            =   2400
               TabIndex        =   218
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   16
               Left            =   1680
               TabIndex        =   217
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   15
               Left            =   840
               TabIndex        =   216
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   215
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   13
               Left            =   840
               TabIndex        =   214
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   213
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   2
               Left            =   840
               TabIndex        =   77
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   76
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   4
               Left            =   1680
               TabIndex        =   75
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   5
               Left            =   2400
               TabIndex        =   74
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Femenino"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   73
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Masculino"
               Height          =   375
               Index           =   7
               Left            =   1680
               TabIndex        =   72
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   71
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   70
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   69
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   840
               TabIndex        =   68
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   67
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   840
               TabIndex        =   66
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   65
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   840
               TabIndex        =   64
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   63
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   9
               Left            =   840
               TabIndex        =   62
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   61
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   11
               Left            =   840
               TabIndex        =   60
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   1680
               TabIndex        =   59
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   2400
               TabIndex        =   58
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   1680
               TabIndex        =   57
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   2400
               TabIndex        =   56
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   1680
               TabIndex        =   55
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   2400
               TabIndex        =   54
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   1680
               TabIndex        =   53
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   2400
               TabIndex        =   52
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   1680
               TabIndex        =   51
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   9
               Left            =   2400
               TabIndex        =   50
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   1680
               TabIndex        =   49
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   11
               Left            =   2400
               TabIndex        =   48
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   1680
               TabIndex        =   47
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   17
               Left            =   2400
               TabIndex        =   46
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   1680
               TabIndex        =   45
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   44
               Top             =   5880
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Módulos"
            Height          =   5415
            Index           =   2
            Left            =   -74880
            TabIndex        =   36
            Top             =   360
            Width           =   1815
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cardiología"
               Height          =   375
               Index           =   45
               Left            =   0
               TabIndex        =   212
               Top             =   5040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tuberculosis"
               Height          =   375
               Index           =   44
               Left            =   0
               TabIndex        =   211
               Top             =   4560
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Somatometría"
               Height          =   375
               Index           =   52
               Left            =   0
               TabIndex        =   42
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laboratorio"
               Height          =   375
               Index           =   51
               Left            =   0
               TabIndex        =   41
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dental"
               Height          =   375
               Index           =   50
               Left            =   0
               TabIndex        =   40
               Top             =   2640
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nutrición"
               Height          =   375
               Index           =   49
               Left            =   0
               TabIndex        =   39
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Optometría"
               Height          =   375
               Index           =   47
               Left            =   0
               TabIndex        =   38
               Top             =   3600
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Audiometría"
               Height          =   375
               Index           =   46
               Left            =   0
               TabIndex        =   37
               Top             =   4080
               Width           =   1815
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   7215
            Index           =   0
            Left            =   1920
            TabIndex        =   13
            Top             =   480
            Width           =   4695
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   23
               Left            =   840
               TabIndex        =   210
               Top             =   6840
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   22
               Left            =   120
               TabIndex        =   209
               Top             =   6840
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   21
               Left            =   840
               TabIndex        =   208
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   20
               Left            =   120
               TabIndex        =   207
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "No."
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   34
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   33
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   32
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   31
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   840
               TabIndex        =   30
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   29
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   840
               TabIndex        =   28
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   27
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   840
               TabIndex        =   26
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   25
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   9
               Left            =   840
               TabIndex        =   24
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   23
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   11
               Left            =   840
               TabIndex        =   22
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   21
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   13
               Left            =   840
               TabIndex        =   20
               Top             =   4440
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   19
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   15
               Left            =   840
               TabIndex        =   18
               Top             =   4920
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   16
               Left            =   120
               TabIndex        =   17
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   17
               Left            =   840
               TabIndex        =   16
               Top             =   5400
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   18
               Left            =   120
               TabIndex        =   15
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H80000007&
               BackStyle       =   0  'Transparent
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   5
               EndProperty
               Height          =   375
               Index           =   19
               Left            =   840
               TabIndex        =   14
               Top             =   5880
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Módulos"
            Height          =   7215
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1815
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cardiología"
               Height          =   375
               Index           =   43
               Left            =   0
               TabIndex        =   206
               Top             =   6840
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tuberculosis"
               Height          =   375
               Index           =   27
               Left            =   0
               TabIndex        =   205
               Top             =   6360
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Somatometría"
               Height          =   375
               Index           =   42
               Left            =   0
               TabIndex        =   12
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laboratorio"
               Height          =   375
               Index           =   41
               Left            =   0
               TabIndex        =   11
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Dental"
               Height          =   375
               Index           =   40
               Left            =   0
               TabIndex        =   10
               Top             =   2520
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nutrición"
               Height          =   375
               Index           =   39
               Left            =   0
               TabIndex        =   9
               Top             =   3000
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Salud de la mujer"
               Height          =   375
               Index           =   38
               Left            =   0
               TabIndex        =   8
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Optometría"
               Height          =   375
               Index           =   37
               Left            =   0
               TabIndex        =   7
               Top             =   5400
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Audiometría"
               Height          =   375
               Index           =   36
               Left            =   0
               TabIndex        =   6
               Top             =   5880
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DOCMA"
               Height          =   375
               Index           =   35
               Left            =   1080
               TabIndex        =   5
               Top             =   3960
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DOCCU"
               Height          =   375
               Index           =   34
               Left            =   1080
               TabIndex        =   4
               Top             =   4440
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mastografía"
               Height          =   375
               Index           =   33
               Left            =   360
               TabIndex        =   3
               Top             =   4920
               Width           =   1455
            End
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    RsSomatometria.Filter = ""
    With RsSomatometria
        If .State = 1 Then .Close
            .Open "Select * from SOMAT", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(0) = RsSomatometria.RecordCount
    Label2(1) = Round(((Label2(0) / Label2(0)) * 100), 2)
    With RsLaboratorio
        If .State = 1 Then .Close
            .Open "Select * from LABOR", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(2) = RsLaboratorio.RecordCount
    Label2(3) = Round(((Label2(2) / Label2(0)) * 100), 2)
    With RsDental
        If .State = 1 Then .Close
            .Open "Select * from DENTA", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(4) = RsDental.RecordCount
    Label2(5) = Round(((Label2(4) / Label2(0)) * 100), 2)
    With RsNutricion
        If .State = 1 Then .Close
            .Open "Select * from NUTRI", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(6) = RsNutricion.RecordCount
    Label2(7) = Round(((Label2(6) / Label2(0)) * 100), 2)
    With RsSaludMujer
        If .State = 1 Then .Close
            .Open "Select * from SD_MU", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    With RsMujer
        If .State = 1 Then .Close
            .Open "Select ID_AST, NOMBRE from SOMAT Where GENERO = 'Femenino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(8) = RsMujer.RecordCount
    Label2(9) = Round(((Label2(8) / Label2(8)) * 100), 2)
    With RsDocma
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ Where COCAMA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(10) = RsDocma.RecordCount
    Label2(11) = Round(((Label2(10) / Label2(8)) * 100), 2)
    With RsDoccu
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ Where COCCU = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(12) = RsDoccu.RecordCount
    Label2(13) = Round(((Label2(12) / Label2(8)) * 100), 2)
    With RsMastografia
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ Where MSTGFA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(14) = RsMastografia.RecordCount
    Label2(15) = Round(((Label2(14) / Label2(8)) * 100), 2)
    With RSOptometria
        If .State = 1 Then .Close
            .Open "Select * from V_OYA Where OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(16) = RSOptometria.RecordCount
    Label2(17) = Round(((Label2(16) / Label2(0)) * 100), 2)
    With RsAudiometria
        If .State = 1 Then .Close
            .Open "Select * from V_AUD Where ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(18) = RsAudiometria.RecordCount
    Label2(19) = Round(((Label2(18) / Label2(0)) * 100), 2)
    With RSATuberculosis
        If .State = 1 Then .Close
            .Open "Select * from V_TUB Where ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(20) = RSATuberculosis.RecordCount
    Label2(21) = Round(((Label2(20) / Label2(0)) * 100), 2)
    With RSACardiologia
        If .State = 1 Then .Close
            .Open "Select * from V_CAR Where ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label2(22) = RSACardiologia.RecordCount
    Label2(23) = Round(((Label2(22) / Label2(0)) * 100), 2)
    With RsGSomatometria
        If .State = 1 Then .Close
            .Open "Select * from SOMAT Where GENERO = 'Femenino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(0) = RsGSomatometria.RecordCount
    Label3(1) = Round(((Label3(0) / Label2(0)) * 100), 2)
    With RsGLaboratorio
        If .State = 1 Then .Close
            .Open "Select * from V_LAB Where GENERO = 'Femenino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(2) = RsGLaboratorio.RecordCount
    Label3(3) = Round(((Label3(2) / Label2(0)) * 100), 2)
    With RsGDental
        If .State = 1 Then .Close
            .Open "Select * from V_DEN Where GENERO = 'Femenino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(4) = RsGDental.RecordCount
    Label3(5) = Round(((Label3(4) / Label2(0)) * 100), 2)
    With RsGNutricion
        If .State = 1 Then .Close
            .Open "Select * from V_NUT Where GENERO = 'Femenino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(6) = RsGNutricion.RecordCount
    Label3(7) = Round(((Label3(6) / Label2(0)) * 100), 2)
    With RSGOptometria
        If .State = 1 Then .Close
            .Open "Select * from V_OYA Where GENERO = 'Femenino' AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(8) = RSGOptometria.RecordCount
    Label3(9) = Round(((Label3(8) / Label2(0)) * 100), 2)
    With RsGAudiometria
        If .State = 1 Then .Close
            .Open "Select * from V_AUD Where GENERO = 'Femenino' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(10) = RsGAudiometria.RecordCount
    Label3(11) = Round(((Label3(10) / Label2(0)) * 100), 2)
    With RsGTuberculosis
        If .State = 1 Then .Close
            .Open "Select * from V_TUB Where GENERO = 'Femenino' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(12) = RsGTuberculosis.RecordCount
    Label3(13) = Round(((Label3(12) / Label2(0)) * 100), 2)
    With RSGCardiologia
        If .State = 1 Then .Close
            .Open "Select * from V_CAR Where GENERO = 'Femenino' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(14) = RSGCardiologia.RecordCount
    Label3(15) = Round(((Label3(14) / Label2(0)) * 100), 2)
    With RsGMSomatometria
        If .State = 1 Then .Close
            .Open "Select * from SOMAT Where GENERO = 'Masculino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label4(0) = RsGMSomatometria.RecordCount
    Label4(1) = Round(((Label4(0) / Label2(0)) * 100), 2)
    With RsGMLaboratorio
        If .State = 1 Then .Close
            .Open "Select * from V_LAB Where GENERO = 'Masculino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label4(2) = RsGMLaboratorio.RecordCount
    Label4(3) = Round(((Label4(2) / Label2(0)) * 100), 2)
    With RsGMDental
        If .State = 1 Then .Close
            .Open "Select * from V_DEN Where GENERO = 'Masculino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label4(4) = RsGMDental.RecordCount
    Label4(5) = Round(((Label4(4) / Label2(0)) * 100), 2)
    With RsGMNutricion
        If .State = 1 Then .Close
            .Open "Select * from V_NUT Where GENERO = 'Masculino'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label4(6) = RsGMNutricion.RecordCount
    Label4(7) = Round(((Label4(6) / Label2(0)) * 100), 2)
    With RSGMOptometria
        If .State = 1 Then .Close
            .Open "Select * from V_OYA Where GENERO = 'Masculino' AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label4(8) = RSGMOptometria.RecordCount
    Label4(9) = Round(((Label4(8) / Label2(0)) * 100), 2)
    With RsGMAudiometria
        If .State = 1 Then .Close
            .Open "Select * from V_AUD Where GENERO = 'Masculino' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label4(10) = RsGMAudiometria.RecordCount
    Label4(11) = Round(((Label4(10) / Label2(0)) * 100), 2)
    With RSGMTuberculosis
        If .State = 1 Then .Close
            .Open "Select * from V_TUB Where GENERO = 'Masculino' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(16) = RSGMTuberculosis.RecordCount
    Label3(17) = Round(((Label3(16) / Label2(0)) * 100), 2)
    With RSGMCardiologia
        If .State = 1 Then .Close
            .Open "Select * from V_CAR Where GENERO = 'Masculino' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label3(18) = RSGMCardiologia.RecordCount
    Label3(19) = Round(((Label3(18) / Label2(0)) * 100), 2)
    With ESO
        If .State = 1 Then .Close
            .Open "Select * from SOMAT where edad <31", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(0) = ESO.RecordCount
    Label5(1) = Round(((Label5(0) / Label2(0)) * 100), 2)
    With ELO
        If .State = 1 Then .Close
            .Open "Select * from V_LAB where edad <31", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(2) = ELO.RecordCount
    Label5(3) = Round(((Label5(2) / Label2(0)) * 100), 2)
    With EDO
        If .State = 1 Then .Close
            .Open "Select * from V_DEN where edad <31", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(4) = EDO.RecordCount
    Label5(5) = Round(((Label5(4) / Label2(0)) * 100), 2)
    With ENO
        If .State = 1 Then .Close
            .Open "Select * from V_NUT Where EDAD <31", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(6) = ENO.RecordCount
    Label5(7) = Round(((Label5(6) / Label2(0)) * 100), 2)
    With EDMO
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad <31 AND COCAMA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(8) = EDMO.RecordCount
    Label5(9) = Round(((Label5(8) / Label2(8)) * 100), 2)
    With EDCO
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad <31 AND COCCU = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(10) = EDCO.RecordCount
    Label5(11) = Round(((Label5(10) / Label2(8)) * 100), 2)
    With EMGO
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad <31 AND MSTGFA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(12) = EMGO.RecordCount
    Label5(13) = Round(((Label5(12) / Label2(8)) * 100), 2)
    With EOO
        If .State = 1 Then .Close
            .Open "Select * from V_OYA where edad <31 AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(14) = EOO.RecordCount
    Label5(15) = Round(((Label5(14) / Label2(0)) * 100), 2)
    With EAO
        If .State = 1 Then .Close
            .Open "Select * from V_AUD where edad <31 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(16) = EAO.RecordCount
    Label5(17) = Round(((Label5(16) / Label2(0)) * 100), 2)
    With ETO
        If .State = 1 Then .Close
            .Open "Select * from V_TUB where edad <31 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(18) = ETO.RecordCount
    Label5(19) = Round(((Label5(18) / Label2(0)) * 100), 2)
    With ECO
        If .State = 1 Then .Close
            .Open "Select * from V_CAR where edad <31 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Label5(20) = ECO.RecordCount
    Label5(21) = Round(((Label5(20) / Label2(0)) * 100), 2)
    With ES3O
        If .State = 1 Then .Close
            .Open "Select * from SOMAT where edad >30 AND EDAD <61", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(0) = ES3O.RecordCount
    Labe6(1) = Round(((Labe6(0) / Label2(0)) * 100), 2)
    With EL3O
        If .State = 1 Then .Close
            .Open "Select * from V_LAB where edad >30 AND EDAD <61", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(2) = EL3O.RecordCount
    Labe6(3) = Round(((Labe6(2) / Label2(0)) * 100), 2)
    With ED3O
        If .State = 1 Then .Close
            .Open "Select * from V_DEN where edad >30 AND EDAD <61", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(4) = ED3O.RecordCount
    Labe6(5) = Round(((Labe6(4) / Label2(0)) * 100), 2)
    With EN3O
        If .State = 1 Then .Close
            .Open "Select * from V_NUT Where EDAD >30 AND EDAD <61", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(6) = EN3O.RecordCount
    Labe6(7) = Round(((Labe6(6) / Label2(0)) * 100), 2)
    With EDM3O
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad >30 AND EDAD <61 AND COCAMA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(8) = EDM3O.RecordCount
    Labe6(9) = Round(((Labe6(8) / Label2(8)) * 100), 2)
    With EDC3O
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad >30 AND EDAD <61 AND COCCU = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(10) = EDC3O.RecordCount
    Labe6(11) = Round(((Labe6(10) / Label2(8)) * 100), 2)
    With EMG3O
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad >30 AND EDAD <61 AND MSTGFA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(12) = EMG3O.RecordCount
    Labe6(13) = Round(((Labe6(12) / Label2(8)) * 100), 2)
    With EO3O
        If .State = 1 Then .Close
            .Open "Select * from V_OYA where edad >30 AND EDAD <61 AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(14) = EO3O.RecordCount
    Labe6(15) = Round(((Labe6(14) / Label2(0)) * 100), 2)
    With EA3O
        If .State = 1 Then .Close
            .Open "Select * from V_AUD where edad >30 AND EDAD <61 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(16) = EA3O.RecordCount
    Labe6(17) = Round(((Labe6(16) / Label2(0)) * 100), 2)
    With ET3O
        If .State = 1 Then .Close
            .Open "Select * from V_TUB where edad >30 AND EDAD <61 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(18) = ET3O.RecordCount
    Labe6(19) = Round(((Labe6(18) / Label2(0)) * 100), 2)
    With EC3O
        If .State = 1 Then .Close
            .Open "Select * from V_CAR where edad >30 AND EDAD <61 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe6(20) = EC3O.RecordCount
    Labe6(21) = Round(((Labe6(20) / Label2(0)) * 100), 2)
    With ES6O
        If .State = 1 Then .Close
            .Open "Select * from SOMAT where edad >60", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(0) = ES6O.RecordCount
    Labe7(1) = Round(((Labe7(0) / Label2(0)) * 100), 2)
    With EL6O
        If .State = 1 Then .Close
            .Open "Select * from V_LAB where edad >60", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(2) = EL6O.RecordCount
    Labe7(3) = Round(((Labe7(2) / Label2(0)) * 100), 2)
    With ED6O
        If .State = 1 Then .Close
            .Open "Select * from V_DEN where edad >60", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(4) = ED6O.RecordCount
    Labe7(5) = Round(((Labe7(4) / Label2(0)) * 100), 2)
    With EN6O
        If .State = 1 Then .Close
            .Open "Select * from V_NUT Where EDAD >60", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(6) = EN6O.RecordCount
    Labe7(7) = Round(((Labe7(6) / Label2(0)) * 100), 2)
    With EDM6O
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad >60 AND COCAMA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(8) = EDM6O.RecordCount
    Labe7(9) = Round(((Labe7(8) / Label2(8)) * 100), 2)
    With EDC6O
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad >60 AND COCCU = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(10) = EDC6O.RecordCount
    Labe7(11) = Round(((Labe7(10) / Label2(8)) * 100), 2)
    With EMG6O
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where edad >60 AND MSTGFA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(12) = EMG6O.RecordCount
    Labe7(13) = Round(((Labe7(12) / Label2(8)) * 100), 2)
    With EO6O
        If .State = 1 Then .Close
            .Open "Select * from V_OYA where edad >60 AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(14) = EO6O.RecordCount
    Labe7(15) = Round(((Labe7(14) / Label2(0)) * 100), 2)
    With EA6O
        If .State = 1 Then .Close
            .Open "Select * from V_AUD where edad >60 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(16) = EA6O.RecordCount
    Labe7(17) = Round(((Labe7(16) / Label2(0)) * 100), 2)
    With ET6O
        If .State = 1 Then .Close
            .Open "Select * from V_TUB where edad >60 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(18) = ET6O.RecordCount
    Labe7(19) = Round(((Labe7(18) / Label2(0)) * 100), 2)
    With EC6O
        If .State = 1 Then .Close
            .Open "Select * from V_CAR where edad >60 AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe7(20) = EC6O.RecordCount
    Labe7(21) = Round(((Labe7(20) / Label2(0)) * 100), 2)
    With TSOM
        If .State = 1 Then .Close
            .Open "Select * from SOMAT where TRAB_E = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(0) = TSOM.RecordCount
    Labe8(1) = Round(((Labe8(0) / Label2(0)) * 100), 2)
    With TLAB
        If .State = 1 Then .Close
            .Open "Select * from V_LAB where TRAB_E = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(2) = TLAB.RecordCount
    Labe8(3) = Round(((Labe8(2) / Label2(0)) * 100), 2)
    With TDEN
        If .State = 1 Then .Close
            .Open "Select * from V_DEN where TRAB_E = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(4) = TDEN.RecordCount
    Labe8(5) = Round(((Labe8(4) / Label2(0)) * 100), 2)
    With TNUT
        If .State = 1 Then .Close
            .Open "Select * from V_NUT where TRAB_E = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(6) = TNUT.RecordCount
    Labe8(7) = Round(((Labe8(6) / Label2(0)) * 100), 2)
    With TCMA
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where TRAB_E = 'Si' AND COCAMA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(8) = TCMA.RecordCount
    Labe8(9) = Round(((Labe8(8) / Label2(8)) * 100), 2)
    With TDOC
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where TRAB_E = 'Si' AND COCCU = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(10) = TDOC.RecordCount
    Labe8(11) = Round(((Labe8(10) / Label2(8)) * 100), 2)
    With TMAS
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where TRAB_E = 'Si' AND MSTGFA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(12) = TMAS.RecordCount
    Labe8(13) = Round(((Labe8(12) / Label2(8)) * 100), 2)
    With TOPT
        If .State = 1 Then .Close
            .Open "Select * from V_OYA where TRAB_E = 'Si' AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(14) = TOPT.RecordCount
    Labe8(15) = Round(((Labe8(14) / Label2(0)) * 100), 2)
    With TAUD
        If .State = 1 Then .Close
            .Open "Select * from V_AUD where TRAB_E = 'Si' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(16) = TAUD.RecordCount
    Labe8(17) = Round(((Labe8(16) / Label2(0)) * 100), 2)
    With TTUB
        If .State = 1 Then .Close
            .Open "Select * from V_TUB where TRAB_E = 'Si' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(18) = TTUB.RecordCount
    Labe8(19) = Round(((Labe8(18) / Label2(0)) * 100), 2)
    With TCAR
        If .State = 1 Then .Close
            .Open "Select * from V_CAR where TRAB_E = 'Si' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe8(20) = TCAR.RecordCount
    Labe8(21) = Round(((Labe8(20) / Label2(0)) * 100), 2)
    With ISOM
        If .State = 1 Then .Close
            .Open "Select * from SOMAT where TRAB_E = 'No'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(0) = ISOM.RecordCount
    Labe9(1) = Round(((Labe9(0) / Label2(0)) * 100), 2)
    With ILAB
        If .State = 1 Then .Close
            .Open "Select * from V_LAB where TRAB_E = 'No'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(2) = ILAB.RecordCount
    Labe9(3) = Round(((Labe9(2) / Label2(0)) * 100), 2)
    With IDEN
        If .State = 1 Then .Close
            .Open "Select * from V_DEN where TRAB_E = 'No'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(4) = IDEN.RecordCount
    Labe9(5) = Round(((Labe9(4) / Label2(0)) * 100), 2)
    With INUT
        If .State = 1 Then .Close
            .Open "Select * from V_NUT where TRAB_E = 'No'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(6) = INUT.RecordCount
    Labe9(7) = Round(((Labe9(6) / Label2(0)) * 100), 2)
    With ICMA
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where TRAB_E = 'No' AND COCAMA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(8) = ICMA.RecordCount
    Labe9(9) = Round(((Labe9(8) / Label2(8)) * 100), 2)
    With IDOC
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where TRAB_E = 'No' AND COCCU = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(10) = IDOC.RecordCount
    Labe9(11) = Round(((Labe9(10) / Label2(8)) * 100), 2)
    With IMAS
        If .State = 1 Then .Close
            .Open "Select * from V_MUJ where TRAB_E = 'No' AND MSTGFA = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(12) = IMAS.RecordCount
    Labe9(13) = Round(((Labe9(12) / Label2(8)) * 100), 2)
    With IOPT
        If .State = 1 Then .Close
            .Open "Select * from V_OYA where TRAB_E = 'No' AND OPTOME = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(14) = IOPT.RecordCount
    Labe9(15) = Round(((Labe9(14) / Label2(0)) * 100), 2)
    With IAUD
        If .State = 1 Then .Close
            .Open "Select * from V_AUD where TRAB_E = 'No' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(16) = IAUD.RecordCount
    Labe9(17) = Round(((Labe9(16) / Label2(0)) * 100), 2)
    With ITUB
        If .State = 1 Then .Close
            .Open "Select * from V_TUB where TRAB_E = 'No' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(18) = ITUB.RecordCount
    Labe9(19) = Round(((Labe9(18) / Label2(0)) * 100), 2)
    With ICAR
        If .State = 1 Then .Close
            .Open "Select * from V_CAR where TRAB_E = 'No' AND ASISTE = 'Si'", Cn, adOpenStatic, adLockOptimistic
            .Requery
    End With
    Labe9(20) = ICAR.RecordCount
    Labe9(21) = Round(((Labe9(20) / Label2(0)) * 100), 2)
End Sub
Private Sub Salir_Click()
    On Error Resume Next
    Form1.Enabled = True
    Unload Me
End Sub
