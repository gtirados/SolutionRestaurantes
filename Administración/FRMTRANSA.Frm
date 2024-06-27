VERSION 5.00
Begin VB.Form FRMTRANS 
   BackColor       =   &H00F5F1EC&
   Caption         =   "TRANS"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   930
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Icon            =   "FRMTRANSA.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox tra_subkey 
      Height          =   285
      Left            =   1440
      TabIndex        =   204
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Desplaza 
      Caption         =   "&Desplazamiento"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   202
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton CmdCopia 
      Caption         =   "Copia de &Registro"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   199
      Top             =   6000
      Width           =   2775
   End
   Begin VB.ListBox GRUPOS 
      Height          =   1620
      Left            =   1500
      TabIndex        =   167
      Top             =   3030
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame F4 
      BackColor       =   &H00F5F1EC&
      Caption         =   "Reporte :"
      Height          =   1335
      Left            =   3480
      TabIndex        =   178
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
      Begin VB.TextBox Text1 
         DataField       =   "TRA_REP3"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   74
         Left            =   360
         MaxLength       =   3
         TabIndex        =   44
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_REP2"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   73
         Left            =   360
         MaxLength       =   3
         TabIndex        =   43
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_REP1"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   72
         Left            =   360
         MaxLength       =   3
         TabIndex        =   42
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "3.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   165
         Index           =   48
         Left            =   120
         TabIndex        =   181
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "2.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   165
         Index           =   47
         Left            =   120
         TabIndex        =   180
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "1.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   165
         Index           =   46
         Left            =   120
         TabIndex        =   179
         Top             =   240
         Width           =   180
      End
   End
   Begin VB.Frame F3 
      BackColor       =   &H00F5F1EC&
      Caption         =   "GRUPO"
      Height          =   2055
      Left            =   0
      TabIndex        =   154
      Top             =   2760
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU10"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   71
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   41
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU9"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   70
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   40
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU8"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   69
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   39
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU7"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   68
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   38
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU6"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   67
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU5"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   66
         Left            =   360
         MaxLength       =   4
         TabIndex        =   36
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU4"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   65
         Left            =   360
         MaxLength       =   4
         TabIndex        =   35
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU3"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   64
         Left            =   360
         MaxLength       =   4
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU2"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   360
         MaxLength       =   4
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_GRU1"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   360
         MaxLength       =   4
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   67
         Left            =   2640
         TabIndex        =   177
         Top             =   360
         Width           =   525
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   68
         Left            =   2640
         TabIndex        =   176
         Top             =   720
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   69
         Left            =   2640
         TabIndex        =   175
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   70
         Left            =   2640
         TabIndex        =   174
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   71
         Left            =   2640
         TabIndex        =   173
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   32
         Left            =   960
         TabIndex        =   172
         Top             =   285
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   33
         Left            =   960
         TabIndex        =   171
         Top             =   645
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   64
         Left            =   960
         TabIndex        =   170
         Top             =   1035
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   65
         Left            =   960
         TabIndex        =   169
         Top             =   1380
         Width           =   510
      End
      Begin VB.Label LBLG 
         BackColor       =   &H00F5F1EC&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   66
         Left            =   960
         TabIndex        =   168
         Top             =   1740
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "10.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   45
         Left            =   1800
         TabIndex        =   164
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "9.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   44
         Left            =   1800
         TabIndex        =   163
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "8.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   43
         Left            =   1800
         TabIndex        =   162
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "7.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   42
         Left            =   1800
         TabIndex        =   161
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "6.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   41
         Left            =   1800
         TabIndex        =   160
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "5.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   40
         Left            =   105
         TabIndex        =   159
         Top             =   1710
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "4.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   39
         Left            =   120
         TabIndex        =   158
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "3.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   38
         Left            =   120
         TabIndex        =   157
         Top             =   1005
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "2.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   156
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "1.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   155
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton CmdOtro 
      Caption         =   "&Otro"
      Height          =   375
      Left            =   240
      TabIndex        =   116
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      DataField       =   "TRA_DESCRIPCION"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox txtkey 
      DataField       =   "TRA_KEY"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      MaxLength       =   4
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   117
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel / Retorna"
      Height          =   375
      Left            =   2520
      TabIndex        =   119
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame F2 
      BackColor       =   &H00F5F1EC&
      Caption         =   "ACTUALIZACIONES"
      Height          =   1335
      Left            =   0
      TabIndex        =   143
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox TxtAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   31
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   30
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   28
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         MaxLength       =   3
         TabIndex        =   27
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT1"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   51
         Left            =   360
         MaxLength       =   3
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT2"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   50
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT3"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   49
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT4"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   48
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT5"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   47
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT6"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   46
         Left            =   360
         MaxLength       =   3
         TabIndex        =   22
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT7"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   45
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT8"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   44
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   24
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT9"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_ACT10"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "15.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   57
         Left            =   3360
         TabIndex        =   189
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "14.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   56
         Left            =   2520
         TabIndex        =   188
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "13.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   55
         Left            =   1680
         TabIndex        =   187
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "12.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   54
         Left            =   840
         TabIndex        =   186
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "11.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   53
         Left            =   0
         TabIndex        =   185
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "1.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   37
         Left            =   120
         TabIndex        =   153
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "2.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   36
         Left            =   960
         TabIndex        =   152
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "3.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   35
         Left            =   1800
         TabIndex        =   151
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "4.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   34
         Left            =   2640
         TabIndex        =   150
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "5.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   33
         Left            =   3480
         TabIndex        =   149
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "6.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   32
         Left            =   120
         TabIndex        =   148
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "7.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   31
         Left            =   960
         TabIndex        =   147
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "8.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   30
         Left            =   1800
         TabIndex        =   146
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "9.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   29
         Left            =   2640
         TabIndex        =   145
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "10.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   28
         Left            =   3360
         TabIndex        =   144
         Top             =   600
         Width           =   345
      End
   End
   Begin VB.Frame F1 
      BackColor       =   &H00F5F1EC&
      Caption         =   "CONSISTENCIAS"
      Height          =   1335
      Left            =   0
      TabIndex        =   130
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox TxtCons 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtCons 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtCons 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON12"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   63
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON8"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   62
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON4"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   61
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON11"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   60
         Left            =   360
         MaxLength       =   3
         TabIndex        =   12
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON7"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   59
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON3"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   58
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON10"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   57
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON6"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   55
         Left            =   360
         MaxLength       =   3
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON2"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   53
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON9"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   56
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON5"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   54
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_CON1"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   52
         Left            =   360
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "15.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   52
         Left            =   3360
         TabIndex        =   184
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "14.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   51
         Left            =   2520
         TabIndex        =   183
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "13.-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   50
         Left            =   1680
         TabIndex        =   182
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "12.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   25
         Left            =   840
         TabIndex        =   142
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "8.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   24
         Left            =   1800
         TabIndex        =   141
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "4.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   23
         Left            =   2610
         TabIndex        =   140
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "11.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   22
         Left            =   0
         TabIndex        =   139
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "7.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   21
         Left            =   960
         TabIndex        =   138
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "3.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   20
         Left            =   1800
         TabIndex        =   137
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "10.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   19
         Left            =   3360
         TabIndex        =   136
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "6.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   135
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "2.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   17
         Left            =   960
         TabIndex        =   134
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "9.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   16
         Left            =   2610
         TabIndex        =   133
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "5.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   13
         Left            =   3450
         TabIndex        =   132
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "1.- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   131
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame F5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Secu.  ------ Campo ----------------  Texto ----------------------  Top ----  Left"
      Height          =   5775
      Left            =   4680
      TabIndex        =   118
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S14"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   90
         Left            =   360
         MaxLength       =   3
         TabIndex        =   203
         Top             =   5400
         Width           =   350
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   80
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   198
         Top             =   5400
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   17
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   197
         Top             =   5400
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   18
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   196
         Top             =   5400
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   16
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   194
         Top             =   5400
         Width           =   375
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S14"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   89
         Left            =   360
         MaxLength       =   3
         TabIndex        =   115
         Top             =   5040
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S13"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   88
         Left            =   360
         MaxLength       =   3
         TabIndex        =   114
         Top             =   4680
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S12"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   87
         Left            =   360
         MaxLength       =   3
         TabIndex        =   113
         Top             =   4320
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S11"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   86
         Left            =   360
         MaxLength       =   3
         TabIndex        =   112
         Top             =   3960
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S10"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   85
         Left            =   360
         MaxLength       =   3
         TabIndex        =   111
         Top             =   3600
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S9"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   84
         Left            =   360
         MaxLength       =   3
         TabIndex        =   110
         Top             =   3240
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S8"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   83
         Left            =   360
         MaxLength       =   3
         TabIndex        =   109
         Top             =   2880
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S7"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   82
         Left            =   360
         MaxLength       =   3
         TabIndex        =   108
         Top             =   2520
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S6"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   81
         Left            =   360
         MaxLength       =   3
         TabIndex        =   107
         Top             =   2160
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S5"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   80
         Left            =   360
         MaxLength       =   3
         TabIndex        =   106
         Top             =   1800
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S4"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   79
         Left            =   360
         MaxLength       =   3
         TabIndex        =   105
         Top             =   1440
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S3"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   78
         Left            =   360
         MaxLength       =   3
         TabIndex        =   104
         Top             =   1080
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   77
         Left            =   360
         MaxLength       =   3
         TabIndex        =   103
         Top             =   720
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_S1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   76
         Left            =   360
         MaxLength       =   3
         TabIndex        =   102
         Top             =   360
         Width           =   345
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   15
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   101
         Top             =   5040
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   14
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   100
         Top             =   5040
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   13
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   99
         Top             =   5040
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   12
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   98
         Top             =   5040
         Width           =   350
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   11
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   97
         Top             =   4680
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   10
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   96
         Top             =   4680
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   95
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   8
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   94
         Top             =   4680
         Width           =   350
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   7
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   93
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   6
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   92
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   5
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   91
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   90
         Top             =   4320
         Width           =   350
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   3
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   89
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   2
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   88
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   87
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox TxtCampos 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   86
         Top             =   3960
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF10"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   43
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   85
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T10"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   42
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   84
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L10"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   41
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   83
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C10"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   40
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   82
         Top             =   3600
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF9"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   39
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   81
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T9"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   38
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   80
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L9"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   37
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   79
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C9"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   36
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   78
         Top             =   3240
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF8"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   31
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   77
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T8"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   30
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   76
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L8"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   29
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   75
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C8"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   28
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   74
         Top             =   2880
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF7"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   27
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   73
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T7"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   26
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   72
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L7"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   25
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   71
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C7"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   24
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   70
         Top             =   2520
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF6"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   23
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   69
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T6"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   22
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   68
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L6"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   21
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   67
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C6"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   20
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   66
         Top             =   2160
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF5"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   19
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   65
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T5"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   18
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   64
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L5"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   17
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   63
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C5"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   16
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   62
         Top             =   1800
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF4"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   15
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   61
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T4"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   14
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   60
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L4"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   13
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   59
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C4"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   12
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   58
         Top             =   1440
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF3"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   11
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   57
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T3"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   10
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   56
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L3"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   55
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C3"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   8
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   54
         Top             =   1080
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   7
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   53
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   6
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   52
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   5
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   51
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C2"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   50
         Top             =   720
         Width           =   350
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_LF1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   49
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_T1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_L1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   47
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "TRA_C1"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   46
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "15 .- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   62
         Left            =   840
         TabIndex        =   195
         Top             =   5400
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "14.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   61
         Left            =   840
         TabIndex        =   193
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "13.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   60
         Left            =   840
         TabIndex        =   192
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "12.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   59
         Left            =   840
         TabIndex        =   191
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "11.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   58
         Left            =   840
         TabIndex        =   190
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "10.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   15
         Left            =   840
         TabIndex        =   129
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "9.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   14
         Left            =   840
         TabIndex        =   128
         Top             =   3240
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "8.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   12
         Left            =   840
         TabIndex        =   127
         Top             =   2880
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "7.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   126
         Top             =   2520
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "6.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   125
         Top             =   2160
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "4.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   8
         Left            =   840
         TabIndex        =   123
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   122
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   120
         Top             =   360
         Width           =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   5
         X1              =   720
         X2              =   960
         Y1              =   2040
         Y2              =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "5.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   9
         Left            =   840
         TabIndex        =   124
         Top             =   1800
         Width           =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   3
         X1              =   720
         X2              =   960
         Y1              =   1320
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   4
         X1              =   720
         X2              =   960
         Y1              =   1680
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   2
         X1              =   720
         X2              =   960
         Y1              =   960
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2.- "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   121
         Top             =   720
         Width           =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   1
         X1              =   720
         X2              =   960
         Y1              =   600
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   6
         X1              =   720
         X2              =   960
         Y1              =   2400
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   7
         X1              =   720
         X2              =   960
         Y1              =   2760
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   8
         X1              =   720
         X2              =   960
         Y1              =   3120
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   9
         X1              =   720
         X2              =   960
         Y1              =   3480
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   10
         X1              =   720
         X2              =   960
         Y1              =   3840
         Y2              =   4080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   11
         X1              =   720
         X2              =   960
         Y1              =   4200
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   12
         X1              =   720
         X2              =   960
         Y1              =   4560
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H009C3000&
         BorderWidth     =   2
         Index           =   13
         X1              =   720
         X2              =   960
         Y1              =   4920
         Y2              =   5160
      End
   End
   Begin VB.Frame F6 
      BackColor       =   &H00F5F1EC&
      Height          =   735
      Left            =   3480
      TabIndex        =   200
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
      Begin VB.TextBox TRA_FLAG_EXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         MaxLength       =   1
         TabIndex        =   45
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F1EC&
         Caption         =   "N, X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009C3000&
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   201
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "2280"
      Height          =   255
      Index           =   9
      Left            =   5160
      TabIndex        =   215
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "1700"
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   214
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "1300"
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   213
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "5700"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   212
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "7600"
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   211
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "8500"
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   210
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "500"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   209
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "900"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   208
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "3100"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   207
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F5F1EC&
      Caption         =   "1080"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   206
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F5F1EC&
      Caption         =   "Codigo Auxiliar :"
      Height          =   255
      Left            =   0
      TabIndex        =   205
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descrip:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   166
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " KEY :"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   165
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "FRMTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim G1 As Integer
Dim IND As Integer
Dim PasoD As Integer

Public Sub MUESTRA_GRUPOS()
Dim cad
If G1 = 1 Then
   GoTo FIN
End If
G1 = 1
PUB_TIPREG = 80
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
GRUPOS.ToolTipText = "TAB_TIPREG = 80"
GRUPOS.Clear
Do Until tab_mayor.EOF
    GRUPOS.AddItem tab_mayor!TAB_NOMLARGO & String(50, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
    tab_mayor.MoveNext
Loop
Exit Sub
FIN:
End Sub

Public Sub BUSCA(VAR1 As Integer)
If Text1(VAR1).text = "" Then
  LBLG(VAR1).Caption = ""
  Exit Sub
End If
For fila = 0 To GRUPOS.ListCount - 1
 GRUPOS.ListIndex = fila
 If Trim(Right(GRUPOS.text, 4)) = Text1(VAR1).text Then
    LBLG(VAR1).Caption = Left(GRUPOS.text, 40)
    Exit Sub
  End If
Next fila
End Sub
Public Sub VERIFICA()
Dim i As Integer
For i = 0 To 89
If i <> 75 Then
  If Text1(i).text = "" Or Left(Text1(i).text, 1) = "0" Then
    If i = 1 Or i = 5 Or i = 9 Or i = 13 Or i = 17 Or i = 21 Or i = 25 Or i = 29 Or i = 37 Or i = 41 Then
      Text1(i) = " "
    Else
      Text1(i) = 0
    End If
  End If
End If
Next i
For i = 0 To 16
 If TxtCampos(i).text = "" Or Left(TxtCampos(i).text, 1) = "0" Then
   If i = 1 Or i = 5 Or i = 9 Or i = 13 Then
    TxtCampos(i) = " "
   Else
     TxtCampos(i) = 0
   End If
 End If
Next i
For i = 0 To 4
 If TxtAct(i).text = "" Then
   TxtAct(i) = 0
 End If
Next i

For i = 0 To 2
 If TxtCons(i).text = "" Then
   TxtCons(i) = 0
 End If
Next i
If TRA_FLAG_EXT.text = "" Then
   TRA_FLAG_EXT.text = " "
End If

End Sub


Private Sub cmdClose_Click()
FRM_STATUS = "0"
Unload FRMTRANS
FrmTabla1.Show

End Sub





Private Sub CmdCopia_Click()
Dim ww1 As Integer
 pub_mensaje = " Introduzca la Nueva Transacción, debe ser numerico de 4 digitos :"
 Pub_Titulo = "Copia de Transacción "
 mivalor = InputBox(Mensaje, Título, valorpred)
 If mivalor = "" Then
    Exit Sub
 End If
 If Len(Trim(mivalor)) > 4 Or Len(Trim(mivalor)) < 4 Or Len(Trim(mivalor)) = 0 Then
    MsgBox "Son solo 4 digitos .... Reintente hacer la copia !!!", 48, Pub_Titulo
    Exit Sub
 End If
 If InStr(1, mivalor, ".") > 0 Then
    MsgBox "Debe ser Entero  .... Reintente hacer la copia !!!", 48, Pub_Titulo
    Exit Sub
 End If
 If Not IsNumeric(mivalor) Then
    MsgBox "Debe ser Numerico .... Reintente hacer la copia !!!", 48, Pub_Titulo
    Exit Sub
 End If
SQ_OPER = 1
PUB_CODTRA = Val(mivalor)
LEER_TRA_LLAVE
If Not tra_llave.EOF Then
   MsgBox "KEY de Transacción ya EXISTE ... Reintente hacer la copia !!!", 48, Pub_Titulo
   Exit Sub
End If
 VERIFICA
 Gen_llave.AddNew
  Gen_llave!TRA_KEY = Val(mivalor)
  Gen_llave!tra_descripcion = Text3.text
  ww1 = 0
   For ww1 = 0 To 31
      If Gen_llave(2 + ww1).Type = 2 Or Gen_llave(2 + ww1).Type = 3 Or Gen_llave(2 + ww1).Type = 4 Then
        Gen_llave(2 + ww1) = Val(Text1(ww1).text)
      Else
        Gen_llave(2 + ww1) = Text1(ww1).text
      End If
   Next ww1
  For ww1 = 76 To 89
      If Gen_llave(ww1 + 31).Type = 2 Or Gen_llave(ww1 + 31).Type = 3 Or Gen_llave(ww1 + 31).Type = 4 Then
        Gen_llave(ww1 + 31) = Val(Text1(ww1).text)
      Else
        Gen_llave(ww1 + 31) = Text1(ww1).text
      End If
   Next ww1

   For ww1 = 42 To 42 + 16
     If Gen_llave(ww1).Type = 2 Or Gen_llave(ww1).Type = 3 Or Gen_llave(ww1).Type = 4 Then
        Gen_llave(ww1) = Val(TxtCampos(ww1 - 42).text)
      Else
        Gen_llave(ww1) = TxtCampos(ww1 - 42).text
      End If
  Next ww1
  Gen_llave!TRA_C9 = Val(Text1(36).text)
  Gen_llave!TRA_L9 = Text1(37).text
  Gen_llave!TRA_T9 = Val(Text1(38).text)
  Gen_llave!TRA_LF9 = Val(Text1(39).text)
  Gen_llave!TRA_C10 = Val(Text1(40).text)
  Gen_llave!TRA_L10 = Text1(41).text
  Gen_llave!TRA_T10 = Val(Text1(42).text)
  Gen_llave!TRA_LF10 = Val(Text1(43).text)
  Gen_llave!TRA_T15 = Val(TxtCampos(17).text)
  Gen_llave!TRA_LF15 = Val(TxtCampos(18).text)
  Gen_llave!TRA_L15 = TxtCampos(80).text
  
  Gen_llave!tra_con1 = Val(Text1(52).text)
  Gen_llave!TRA_CON2 = Val(Text1(53).text)
  Gen_llave!TRA_CON3 = Val(Text1(58).text)
  Gen_llave!TRA_CON4 = Val(Text1(61).text)
  Gen_llave!TRA_CON5 = Val(Text1(54).text)
  Gen_llave!tra_con6 = Val(Text1(55).text)
  Gen_llave!tra_con7 = Val(Text1(59).text)
  Gen_llave!TRA_CON8 = Val(Text1(62).text)
  Gen_llave!tra_con9 = Val(Text1(56).text)
  Gen_llave!tra_con10 = Val(Text1(57).text)
  Gen_llave!TRA_CON11 = Val(Text1(60).text)
  Gen_llave!tra_con12 = Val(Text1(63).text)
  Gen_llave!TRA_CON13 = Val(TxtCons(0).text)
  Gen_llave!TRA_CON14 = Val(TxtCons(1).text)
  Gen_llave!TRA_CON15 = Val(TxtCons(2).text)
   Gen_llave!tra_ACT1 = Val(Text1(51).text)
   Gen_llave!tra_ACT2 = Val(Text1(50).text)
   Gen_llave!TRA_ACT3 = Val(Text1(49).text)
   Gen_llave!TRA_ACT4 = Val(Text1(48).text)
   Gen_llave!TRA_ACT5 = Val(Text1(47).text)
   Gen_llave!TRA_ACT6 = Val(Text1(46).text)
   Gen_llave!TRA_ACT7 = Val(Text1(45).text)
   Gen_llave!TRA_ACT8 = Val(Text1(44).text)
   Gen_llave!TRA_ACT9 = Val(Text1(35).text)
   Gen_llave!tra_ACT10 = Val(Text1(34).text)
   Gen_llave!TRA_ACT11 = Val(TxtAct(0).text)
   Gen_llave!TRA_ACT12 = Val(TxtAct(1).text)
   Gen_llave!TRA_ACT13 = Val(TxtAct(2).text)
   Gen_llave!TRA_ACT14 = Val(TxtAct(3).text)
   Gen_llave!TRA_ACT15 = Val(TxtAct(4).text)
   Gen_llave!TRA_GRU1 = Val(Text1(32).text)
   Gen_llave!TRA_GRU2 = Val(Text1(33).text)
   Gen_llave!TRA_GRU3 = Val(Text1(64).text)
   Gen_llave!TRA_GRU4 = Val(Text1(65).text)
   Gen_llave!TRA_GRU5 = Val(Text1(66).text)
   Gen_llave!TRA_GRU6 = Val(Text1(67).text)
   Gen_llave!TRA_GRU7 = Val(Text1(68).text)
   Gen_llave!TRA_GRU8 = Val(Text1(69).text)
   Gen_llave!TRA_GRU9 = Val(Text1(70).text)
   Gen_llave!TRA_GRU10 = Val(Text1(71).text)
   Gen_llave!tra_rep1 = Val(Text1(72).text)
   Gen_llave!TRA_REP2 = Val(Text1(73).text)
   Gen_llave!TRA_REP3 = Val(Text1(74).text)
 Gen_llave.Update
 MsgBox "OK ... listo la Copia !!!", 48, Pub_Titulo
End Sub

Private Sub CmdOtro_Click()
BORRA_FIELDS 89, Text1
BORRA_FIELDS 16, TxtCampos
BORRA_FIELDS 4, TxtAct
BORRA_FIELDS 2, TxtCons
txtkey.text = ""
Text3.text = ""
TxtSigno.text = ""
txtkey.SetFocus
CmdOtro.Enabled = False
cmdUpdate.Enabled = True

End Sub

Private Sub cmdUpdate_Click()
Dim ww1 As Integer
On Error GoTo wlin
 If txtkey.text = "" Or Text3.text = "" Then
     MsgBox "ERROR    :     Falta LLaves... Ingreselas..!", 48, Pub_Titulo
     txtkey.SetFocus
     Exit Sub
 End If
 If Text1(76).text = "" Then
   pub_mensaje = " ¿Desea la Secuencia Predeterminada... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbYes Then
     SECUENCIA
   End If
  End If
  If OP_FORM = "A" Then
    Gen_llave.AddNew
  ElseIf OP_FORM = "D" Then
    pub_mensaje = " ¿Desea Actualizar los Datos... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbYes Then
       Gen_llave.Edit
    Else
       '*** SALE DE SUB
       Exit Sub
    End If
  End If
  VERIFICA
  Gen_llave!TRA_KEY = Val(txtkey.text)
  Gen_llave!tra_descripcion = Text3.text
  ww1 = 0
   For ww1 = 0 To 31
      If Gen_llave(2 + ww1).Type = 2 Or Gen_llave(2 + ww1).Type = 3 Or Gen_llave(2 + ww1).Type = 4 Then
        Gen_llave(2 + ww1) = Val(Text1(ww1).text)
      Else
        Gen_llave(2 + ww1) = Text1(ww1).text
      End If
   Next ww1
  For ww1 = 76 To 89
      If Gen_llave(ww1 + 31).Type = 2 Or Gen_llave(ww1 + 31).Type = 3 Or Gen_llave(ww1 + 31).Type = 4 Then
        Gen_llave(ww1 + 31) = Val(Text1(ww1).text)
      Else
        Gen_llave(ww1 + 31) = Text1(ww1).text
      End If
   Next ww1

   For ww1 = 42 To 42 + 16
     If Gen_llave(ww1).Type = 2 Or Gen_llave(ww1).Type = 3 Or Gen_llave(ww1).Type = 4 Then
        Gen_llave(ww1) = Val(TxtCampos(ww1 - 42).text)
      Else
        Gen_llave(ww1) = TxtCampos(ww1 - 42).text
      End If
  Next ww1
  Gen_llave!TRA_C9 = Val(Text1(36).text)
  Gen_llave!TRA_L9 = Text1(37).text
  Gen_llave!TRA_T9 = Val(Text1(38).text)
  Gen_llave!TRA_LF9 = Val(Text1(39).text)
  Gen_llave!TRA_C10 = Val(Text1(40).text)
  Gen_llave!TRA_L10 = Text1(41).text
  Gen_llave!TRA_T10 = Val(Text1(42).text)
  Gen_llave!TRA_LF10 = Val(Text1(43).text)
  Gen_llave!TRA_T15 = Val(TxtCampos(17).text)
  Gen_llave!TRA_LF15 = Val(TxtCampos(18).text)
  Gen_llave!TRA_L15 = TxtCampos(80).text
  
  Gen_llave!tra_con1 = Val(Text1(52).text)
  Gen_llave!TRA_CON2 = Val(Text1(53).text)
  Gen_llave!TRA_CON3 = Val(Text1(58).text)
  Gen_llave!TRA_CON4 = Val(Text1(61).text)
  Gen_llave!TRA_CON5 = Val(Text1(54).text)
  Gen_llave!tra_con6 = Val(Text1(55).text)
  Gen_llave!tra_con7 = Val(Text1(59).text)
  Gen_llave!TRA_CON8 = Val(Text1(62).text)
  Gen_llave!tra_con9 = Val(Text1(56).text)
  Gen_llave!tra_con10 = Val(Text1(57).text)
  Gen_llave!TRA_CON11 = Val(Text1(60).text)
  Gen_llave!tra_con12 = Val(Text1(63).text)
  Gen_llave!TRA_CON13 = Val(TxtCons(0).text)
  Gen_llave!TRA_CON14 = Val(TxtCons(1).text)
  Gen_llave!TRA_CON15 = Val(TxtCons(2).text)
   Gen_llave!tra_ACT1 = Val(Text1(51).text)
   Gen_llave!tra_ACT2 = Val(Text1(50).text)
   Gen_llave!TRA_ACT3 = Val(Text1(49).text)
   Gen_llave!TRA_ACT4 = Val(Text1(48).text)
   Gen_llave!TRA_ACT5 = Val(Text1(47).text)
   Gen_llave!TRA_ACT6 = Val(Text1(46).text)
   Gen_llave!TRA_ACT7 = Val(Text1(45).text)
   Gen_llave!TRA_ACT8 = Val(Text1(44).text)
   Gen_llave!TRA_ACT9 = Val(Text1(35).text)
   Gen_llave!tra_ACT10 = Val(Text1(34).text)
   Gen_llave!TRA_ACT11 = Val(TxtAct(0).text)
   Gen_llave!TRA_ACT12 = Val(TxtAct(1).text)
   Gen_llave!TRA_ACT13 = Val(TxtAct(2).text)
   Gen_llave!TRA_ACT14 = Val(TxtAct(3).text)
   Gen_llave!TRA_ACT15 = Val(TxtAct(4).text)
   Gen_llave!TRA_GRU1 = Val(Text1(32).text)
   Gen_llave!TRA_GRU2 = Val(Text1(33).text)
   Gen_llave!TRA_GRU3 = Val(Text1(64).text)
   Gen_llave!TRA_GRU4 = Val(Text1(65).text)
   Gen_llave!TRA_GRU5 = Val(Text1(66).text)
   Gen_llave!TRA_GRU6 = Val(Text1(67).text)
   Gen_llave!TRA_GRU7 = Val(Text1(68).text)
   Gen_llave!TRA_GRU8 = Val(Text1(69).text)
   Gen_llave!TRA_GRU9 = Val(Text1(70).text)
   Gen_llave!TRA_GRU10 = Val(Text1(71).text)
   Gen_llave!tra_rep1 = Val(Text1(72).text)
   Gen_llave!TRA_REP2 = Val(Text1(73).text)
   Gen_llave!TRA_REP3 = Val(Text1(74).text)
   Gen_llave!TRA_FLAG_EXT = Nulo_Valors(TRA_FLAG_EXT.text)
   Gen_llave!TRA_s15 = Val(Text1(90).text)
   Gen_llave!tra_subkey = Trim(tra_subkey.text)
   
If OP_FORM = "A" Then
   Gen_llave.Update
   ACEPTA = "F"
   CmdOtro.Enabled = True
   cmdUpdate.Enabled = False
   ACEPTA = "F"
   PasoD = 1
ElseIf OP_FORM = "D" Then
  Gen_llave.Update
  Screen.MousePointer = 0
  ACEPTA = "F"
 'SALE DEL FORMULARIO
  CmdCopia.Enabled = False
  FRM_STATUS = "0"
  Unload FRMTRANS
  FrmTabla1.Show
End If

Exit Sub
wlin:
  MsgBox Err.Description & "LLamar a Computo .. "
  Screen.MousePointer = 0

End Sub

Private Sub Desplaza_Click()
Dim mivalor As String
mivalor = InputBox("Ingrese la Posicion:", "Transacciones", , 100, 100)
'XXXXXXXXXXXXXX
   For ww1 = 42 To 42 + 16
     If Gen_llave(ww1).Type = 2 Or Gen_llave(ww1).Type = 3 Or Gen_llave(ww1).Type = 4 Then
        Gen_llave(ww1) = Val(TxtCampos(ww1 - 42).text)
      Else
        Gen_llave(ww1) = TxtCampos(ww1 - 42).text
     End If
   Next ww1
End Sub

Private Sub Form_Activate()
Dim W1 As Integer
Select Case OP_FORM
Case "X"
     MsgBox "No a Seleccionado NADA ...LLamar   COMPUTO  >>>> Tabla1", 48, Pub_Titulo
     Unload FRMTRANS
Case "A"
  CmdOtro.Visible = True
  CmdOtro.Enabled = False
  txtkey.Locked = False
Case "D" 'EDIATR REGISTROS
  Screen.MousePointer = 11
  txtkey.Locked = True
  CmdCopia.Enabled = True
  DoEvents
  Gen_llave.AbsolutePosition = Posi_Reg
  txtkey.text = Gen_llave!TRA_KEY
  Text3.text = Gen_llave!tra_descripcion
   For W1 = 0 To 31
    DoEvents
    Text1(W1).text = Trim(Nulo_Valors(Gen_llave(W1 + 2)))
  Next W1
  For W1 = 42 To 42 + 16
    DoEvents
     TxtCampos(W1 - 42).text = Trim(Nulo_Valors(Gen_llave(W1)))
  Next W1
  For W1 = 76 To 89
    DoEvents
    Text1(W1).text = Trim(Nulo_Valors(Gen_llave(W1 + 31)))
  Next W1
  DoEvents
  Text1(90).text = Nulo_Valor0(Gen_llave!TRA_s15)
  Text1(36).text = Trim(Gen_llave!TRA_C9)
  Text1(37).text = Trim(Gen_llave!TRA_L9)
  Text1(38).text = Trim(Gen_llave!TRA_T9)
  Text1(39).text = Trim(Gen_llave!TRA_LF9)
  Text1(40).text = Trim(Gen_llave!TRA_C10)
  Text1(41).text = Trim(Gen_llave!TRA_L10)
  Text1(42).text = Trim(Gen_llave!TRA_T10)
  Text1(43).text = Trim(Gen_llave!TRA_LF10)
  TxtCampos(17).text = Trim(Gen_llave!TRA_T15)
  TxtCampos(18).text = Trim(Gen_llave!TRA_LF15)
  TxtCampos(80).text = Trim(Nulo_Valors(Gen_llave!TRA_L15))
  Text1(52).text = Gen_llave!tra_con1
  Text1(53).text = Gen_llave!TRA_CON2
  Text1(58).text = Gen_llave!TRA_CON3
  Text1(61).text = Gen_llave!TRA_CON4
  Text1(54).text = Gen_llave!TRA_CON5
  Text1(55).text = Gen_llave!tra_con6
  Text1(59).text = Gen_llave!tra_con7
  Text1(62).text = Gen_llave!TRA_CON8
  Text1(56).text = Gen_llave!tra_con9
  Text1(57).text = Gen_llave!tra_con10
  Text1(60).text = Gen_llave!TRA_CON11
  Text1(63).text = Gen_llave!tra_con12
  TxtCons(0).text = Gen_llave!TRA_CON13
  TxtCons(1).text = Gen_llave!TRA_CON14
  TxtCons(2).text = Gen_llave!TRA_CON15
  Text1(51).text = Gen_llave!tra_ACT1
  Text1(50).text = Gen_llave!tra_ACT2
  Text1(49).text = Gen_llave!TRA_ACT3
  Text1(48).text = Gen_llave!TRA_ACT4
  Text1(47).text = Gen_llave!TRA_ACT5
  Text1(46).text = Gen_llave!TRA_ACT6
  Text1(45).text = Gen_llave!TRA_ACT7
  Text1(44).text = Gen_llave!TRA_ACT8
  Text1(35).text = Gen_llave!TRA_ACT9
  Text1(34).text = Gen_llave!tra_ACT10
  TxtAct(0).text = Gen_llave!TRA_ACT11
  TxtAct(1).text = Gen_llave!TRA_ACT12
  TxtAct(2).text = Gen_llave!TRA_ACT13
  TxtAct(3).text = Gen_llave!TRA_ACT14
  TxtAct(4).text = Gen_llave!TRA_ACT15
  Text1(32).text = Gen_llave!TRA_GRU1
  Text1(33).text = Gen_llave!TRA_GRU2
  Text1(64).text = Gen_llave!TRA_GRU3
  Text1(65).text = Gen_llave!TRA_GRU4
  Text1(66).text = Gen_llave!TRA_GRU5
  Text1(67).text = Gen_llave!TRA_GRU6
  Text1(68).text = Gen_llave!TRA_GRU7
  Text1(69).text = Gen_llave!TRA_GRU8
  Text1(70).text = Gen_llave!TRA_GRU9
  Text1(71).text = Gen_llave!TRA_GRU10
  Text1(72).text = Nulo_Valor0(Gen_llave!tra_rep1)
  Text1(73).text = Nulo_Valor0(Gen_llave!TRA_REP2)
  Text1(74).text = Nulo_Valor0(Gen_llave!TRA_REP3)
  TxtCons(0).text = Gen_llave!TRA_CON13
  TxtCons(1).text = Gen_llave!TRA_CON14
  TxtCons(2).text = Gen_llave!TRA_CON15
  TxtAct(0).text = Gen_llave!TRA_ACT11
  TRA_FLAG_EXT.text = Nulo_Valors(Gen_llave!TRA_FLAG_EXT)
  tra_subkey.text = Nulo_Valors(Gen_llave!tra_subkey)
  
'  TxtSigno.text = Nulo_Valor(Gen_llave!tra_signo)
  PasoD = 1
  Screen.MousePointer = 0
Case Else
  MsgBox "No a Seleccionado NADA ...LLamar   COMPUTO  >>>> Tabla1", 48, Pub_Titulo
  Unload FRMTRANS
End Select
DoEvents
MUESTRA_GRUPOS
F1.Visible = True
F2.Visible = True
F3.Visible = True
F4.Visible = True
F5.Visible = True
F6.Visible = True
FRMTRANS.Visible = True

End Sub

Private Sub Form_Load()
G1 = 0
PasoD = 0
End Sub

Private Sub GRUPOS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  GRUPOS.Visible = False
  Text1(IND + 1).SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
  Text1(IND).text = Val(Right(GRUPOS.text, 2))
  LBLG(IND).Caption = Left(GRUPOS.text, 25)
  GRUPOS.Visible = False
  If IND = 33 Then
    Text1(64).SetFocus
  Else
    Text1(IND + 1).SetFocus
  End If
End If
KeyAscii = 0


End Sub

Private Sub Text1_Change(Index As Integer)
If Text1(Index).Index > 63 And Text1(Index).Index < 72 Then
    Call BUSCA(Index)
ElseIf Text1(Index).Index > 31 And Text1(Index).Index < 34 Then
    Call BUSCA(Index)
End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
If Index > 63 And Index < 72 Then

ElseIf Index > 31 And Index < 34 Then

Else
 Call Azul(Text1(Index).text, Text1(Index))

End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim a As Integer

If KeyAscii <> 13 Then
   GoTo FIN
End If
If Text1(Index).Index > 63 And Text1(Index).Index < 72 Then
   IND = Text1(Index).Index
   GRUPOS.Visible = True
   GRUPOS.SetFocus
ElseIf Text1(Index).Index > 31 And Text1(Index).Index < 34 Then
   IND = Text1(Index).Index
   GRUPOS.Visible = True
   GRUPOS.SetFocus
End If
FIN:

End Sub

Private Sub TxtAct_GotFocus(Index As Integer)
If Index >= 0 And Index <= 4 Then
 Call Azul(TxtAct(Index).text, TxtAct(Index))
End If

End Sub


Private Sub TxtCampos_GotFocus(Index As Integer)
If Index >= 0 And Index <= 16 Then
 Call Azul(TxtCampos(Index).text, TxtCampos(Index))
End If

End Sub

Private Sub TxtCons_GotFocus(Index As Integer)
If Index >= 0 And Index <= 3 Then
 Call Azul(TxtCons(Index).text, TxtCons(Index))
End If

End Sub

Public Sub SECUENCIA()
 Text1(76).text = Text1(4)
 Text1(77).text = Text1(8)
 Text1(78).text = Text1(12)
 Text1(79).text = Text1(16)
 Text1(80).text = Text1(20)
 Text1(81).text = Text1(24)
 Text1(82).text = Text1(28)
 Text1(83).text = Text1(36)
 Text1(84).text = Text1(40)
 Text1(85).text = TxtCampos(0)
 Text1(86).text = TxtCampos(4)
 Text1(87).text = TxtCampos(8)
 Text1(88).text = TxtCampos(12)
 Text1(89).text = 0
End Sub
