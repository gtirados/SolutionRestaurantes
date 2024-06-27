VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTabla2 
   Caption         =   "Mantenimiento de Tablas"
   ClientHeight    =   5715
   ClientLeft      =   540
   ClientTop       =   1140
   ClientWidth     =   8880
   ControlBox      =   0   'False
   Icon            =   "FrmTabla2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdOtro 
      Caption         =   "&Otro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancela 
      Caption         =   "&Cancelar/Retornar"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9763
      _Version        =   327680
      TabOrientation  =   2
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   441
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "D2"
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Datos2"
      Tab(0).Control(0).Enabled=   -1  'True
      TabCaption(1)   =   "D1"
      TabPicture(1)   =   "FrmTabla2.frx":0442
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Datos1"
      Tab(1).Control(0).Enabled=   0   'False
      Begin VB.PictureBox Datos2 
         Height          =   5415
         Left            =   -74640
         ScaleHeight     =   5355
         ScaleWidth      =   9075
         TabIndex        =   95
         Top             =   120
         Width           =   9135
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NUM"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   45
            Left            =   1680
            TabIndex        =   140
            Top             =   45
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   46
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   139
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   47
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   138
            Top             =   720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   48
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   137
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   49
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   136
            Top             =   1440
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   50
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   135
            Top             =   1800
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   51
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   134
            Top             =   2160
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   52
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   133
            Top             =   2520
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   53
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   132
            Top             =   2880
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   54
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   131
            Top             =   3240
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   55
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   130
            Top             =   3600
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   56
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   129
            Top             =   3960
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   57
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   128
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   58
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   127
            Top             =   4680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   59
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   126
            Top             =   5040
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   60
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   125
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   61
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   124
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   62
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   123
            Top             =   720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   63
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   122
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   64
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   121
            Top             =   1440
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   65
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   120
            Top             =   1800
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   66
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   119
            Top             =   2160
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   67
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   118
            Top             =   2520
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   68
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   117
            Top             =   2880
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   69
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   116
            Top             =   3240
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   70
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   115
            Top             =   3600
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   71
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   114
            Top             =   3960
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   72
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   113
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   73
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   112
            Top             =   4680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   74
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   111
            Top             =   5040
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   75
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   110
            Top             =   0
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   76
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   109
            Top             =   360
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   77
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   108
            Top             =   720
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   78
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   107
            Top             =   1080
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   79
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   106
            Top             =   1440
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   80
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   105
            Top             =   1800
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   81
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   104
            Top             =   2160
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   82
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   103
            Top             =   2520
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   83
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   102
            Top             =   2880
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   84
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   101
            Top             =   3240
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   85
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   100
            Top             =   3600
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   86
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   99
            Top             =   3960
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   87
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   98
            Top             =   4320
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   88
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   97
            Top             =   4680
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   89
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   96
            Top             =   5040
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   45
            Left            =   0
            TabIndex        =   186
            Top             =   60
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   46
            Left            =   0
            TabIndex        =   185
            Top             =   375
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "        "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   184
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   47
            Left            =   0
            TabIndex        =   183
            Top             =   735
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   48
            Left            =   0
            TabIndex        =   182
            Top             =   1095
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   49
            Left            =   0
            TabIndex        =   181
            Top             =   1455
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   50
            Left            =   0
            TabIndex        =   180
            Top             =   1815
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   51
            Left            =   0
            TabIndex        =   179
            Top             =   2175
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   52
            Left            =   0
            TabIndex        =   178
            Top             =   2535
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   53
            Left            =   0
            TabIndex        =   177
            Top             =   2895
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   54
            Left            =   0
            TabIndex        =   176
            Top             =   3255
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   55
            Left            =   0
            TabIndex        =   175
            Top             =   3615
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   56
            Left            =   0
            TabIndex        =   174
            Top             =   3975
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   57
            Left            =   0
            TabIndex        =   173
            Top             =   4335
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   58
            Left            =   0
            TabIndex        =   172
            Top             =   4695
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   59
            Left            =   0
            TabIndex        =   171
            Top             =   5055
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   60
            Left            =   3240
            TabIndex        =   170
            Top             =   15
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   61
            Left            =   3240
            TabIndex        =   169
            Top             =   375
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   62
            Left            =   3240
            TabIndex        =   168
            Top             =   735
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   63
            Left            =   3240
            TabIndex        =   167
            Top             =   1095
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   64
            Left            =   3240
            TabIndex        =   166
            Top             =   1455
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   65
            Left            =   3240
            TabIndex        =   165
            Top             =   1815
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   66
            Left            =   3240
            TabIndex        =   164
            Top             =   2175
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   67
            Left            =   3240
            TabIndex        =   163
            Top             =   2535
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   68
            Left            =   3240
            TabIndex        =   162
            Top             =   2895
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   69
            Left            =   3240
            TabIndex        =   161
            Top             =   3240
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   70
            Left            =   3240
            TabIndex        =   160
            Top             =   3615
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   71
            Left            =   3240
            TabIndex        =   159
            Top             =   3975
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   72
            Left            =   3240
            TabIndex        =   158
            Top             =   4320
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   73
            Left            =   3240
            TabIndex        =   157
            Top             =   4695
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   74
            Left            =   3240
            TabIndex        =   156
            Top             =   5055
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   75
            Left            =   6360
            TabIndex        =   155
            Top             =   15
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   76
            Left            =   6360
            TabIndex        =   154
            Top             =   375
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   77
            Left            =   6360
            TabIndex        =   153
            Top             =   735
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   78
            Left            =   6360
            TabIndex        =   152
            Top             =   1095
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   79
            Left            =   6360
            TabIndex        =   151
            Top             =   1455
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   80
            Left            =   6360
            TabIndex        =   150
            Top             =   1815
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   81
            Left            =   6360
            TabIndex        =   149
            Top             =   2175
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   82
            Left            =   6360
            TabIndex        =   148
            Top             =   2535
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   83
            Left            =   6360
            TabIndex        =   147
            Top             =   2895
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   84
            Left            =   6360
            TabIndex        =   146
            Top             =   3255
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   85
            Left            =   6360
            TabIndex        =   145
            Top             =   3615
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   86
            Left            =   6360
            TabIndex        =   144
            Top             =   3975
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   87
            Left            =   6360
            TabIndex        =   143
            Top             =   4335
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   88
            Left            =   6360
            TabIndex        =   142
            Top             =   4695
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Index           =   89
            Left            =   6360
            TabIndex        =   141
            Top             =   5055
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.PictureBox Datos1 
         Height          =   5415
         Left            =   360
         ScaleHeight     =   5355
         ScaleWidth      =   9075
         TabIndex        =   4
         Top             =   120
         Width           =   9135
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   44
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   49
            Top             =   5040
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   43
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   48
            Top             =   4680
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   42
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   47
            Top             =   4320
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   41
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   46
            Top             =   3960
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   40
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   45
            Top             =   3600
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   39
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   44
            Top             =   3240
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   38
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   43
            Top             =   2880
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   37
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   42
            Top             =   2520
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   36
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   41
            Top             =   2160
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   35
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   40
            Top             =   1800
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   34
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   39
            Top             =   1440
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   33
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   38
            Top             =   1080
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   32
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   37
            Top             =   720
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   31
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   30
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   29
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   34
            Top             =   5040
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   28
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   33
            Top             =   4680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   27
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   32
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   26
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   31
            Top             =   3960
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   25
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   30
            Top             =   3600
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   24
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   29
            Top             =   3240
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   23
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   28
            Top             =   2880
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   22
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   27
            Top             =   2520
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   21
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   20
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   25
            Top             =   1800
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   19
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1440
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   18
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   23
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   17
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   16
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   15
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   14
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   19
            Top             =   5040
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   13
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   18
            Top             =   4680
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   12
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   17
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   11
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   16
            Top             =   3960
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   10
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   15
            Top             =   3600
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   9
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   14
            Top             =   3240
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   8
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   13
            Top             =   2880
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   7
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   12
            Top             =   2520
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   6
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   11
            Top             =   2160
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   10
            Top             =   1800
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   4
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1440
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   3
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   8
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   7
            Top             =   720
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NOMBRE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtFields 
            DataField       =   "ESP_NUM"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   5
            Top             =   45
            Width           =   1500
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   44
            Left            =   6360
            TabIndex        =   94
            Top             =   5055
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   43
            Left            =   6360
            TabIndex        =   93
            Top             =   4695
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   42
            Left            =   6360
            TabIndex        =   92
            Top             =   4335
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   41
            Left            =   6360
            TabIndex        =   91
            Top             =   3975
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   40
            Left            =   6360
            TabIndex        =   90
            Top             =   3615
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   39
            Left            =   6360
            TabIndex        =   89
            Top             =   3255
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   38
            Left            =   6360
            TabIndex        =   88
            Top             =   2895
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   37
            Left            =   6360
            TabIndex        =   87
            Top             =   2535
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   36
            Left            =   6360
            TabIndex        =   86
            Top             =   2175
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   35
            Left            =   6360
            TabIndex        =   85
            Top             =   1815
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   34
            Left            =   6360
            TabIndex        =   84
            Top             =   1455
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   33
            Left            =   6360
            TabIndex        =   83
            Top             =   1095
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   32
            Left            =   6360
            TabIndex        =   82
            Top             =   735
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   31
            Left            =   6360
            TabIndex        =   81
            Top             =   375
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   30
            Left            =   6360
            TabIndex        =   80
            Top             =   15
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   29
            Left            =   3240
            TabIndex        =   79
            Top             =   5055
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   28
            Left            =   3240
            TabIndex        =   78
            Top             =   4695
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   27
            Left            =   3240
            TabIndex        =   77
            Top             =   4320
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   26
            Left            =   3240
            TabIndex        =   76
            Top             =   3975
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   25
            Left            =   3240
            TabIndex        =   75
            Top             =   3615
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   24
            Left            =   3240
            TabIndex        =   74
            Top             =   3240
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   23
            Left            =   3240
            TabIndex        =   73
            Top             =   2895
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   22
            Left            =   3240
            TabIndex        =   72
            Top             =   2535
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   21
            Left            =   3240
            TabIndex        =   71
            Top             =   2175
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   3240
            TabIndex        =   70
            Top             =   1815
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   19
            Left            =   3240
            TabIndex        =   69
            Top             =   1455
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   18
            Left            =   3240
            TabIndex        =   68
            Top             =   1095
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   17
            Left            =   3240
            TabIndex        =   67
            Top             =   735
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   16
            Left            =   3240
            TabIndex        =   66
            Top             =   375
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   15
            Left            =   3240
            TabIndex        =   65
            Top             =   15
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   14
            Left            =   0
            TabIndex        =   64
            Top             =   5055
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   13
            Left            =   0
            TabIndex        =   63
            Top             =   4695
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   0
            TabIndex        =   62
            Top             =   4335
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   0
            TabIndex        =   61
            Top             =   3975
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   0
            TabIndex        =   60
            Top             =   3615
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   59
            Top             =   3255
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   0
            TabIndex        =   58
            Top             =   2895
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   57
            Top             =   2535
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   56
            Top             =   2175
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   55
            Top             =   1815
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   54
            Top             =   1455
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   53
            Top             =   1095
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   52
            Top             =   735
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   51
            Top             =   375
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "*"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   50
            Top             =   60
            Visible         =   0   'False
            Width           =   60
         End
      End
   End
   Begin VB.Label LblTexto 
      Caption         =   "Label2"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   120
      TabIndex        =   187
      Top             =   5640
      Width           =   3855
   End
End
Attribute VB_Name = "FrmTabla2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAceptar_Click()
On Error GoTo SALE
Dim s1 As Integer
s1 = 0

Select Case OP_FORM
Case "A"
 Cta_Add = Cta_Add + 1
 If txtFields(0).text = "" Or txtFields(1).text = "" Then
   MsgBox "ERROR       : Falta LLaves... Ingreselas !  ", 46, Pub_Titulo
   txtFields(0).SetFocus
   Exit Sub
 End If
 Gen_llave.AddNew
 For s1 = 0 To CAMPOS1 - 1
   If s1 > 89 Then
    MsgBox "Solo se agrego 89 campos", 46, Pub_Titulo
    Exit For
   End If
  If Gen_llave(s1).Type = 1 Then
    Gen_llave(s1) = txtFields(s1).text
   Else
   If Gen_llave(s1).Type = 9 Or Gen_llave(s1).Type = 10 Or Gen_llave(s1).Type = 11 Then
      Gen_llave(s1) = NULO_DATE(txtFields(s1).text)
   Else
    Gen_llave(s1) = Val(txtFields(s1).text)
   End If
   End If
   
 Next s1
Gen_llave.Update
CmdOtro.Enabled = True
CmdAceptar.Enabled = False
ACEPTA = "F"
Case "D"
 Gen_llave.Edit
 For s1 = 0 To CAMPOS1 - 1
   If s1 > 89 Then
    MsgBox "Solo se actualizo 89 campos", 46, Pub_Titulo
    Exit For
   End If
   If Gen_llave(s1).Type = 1 Then
    Gen_llave(s1) = Nulo_Valors(txtFields(s1).text)
   Else
   If Gen_llave(s1).Type = 9 Or Gen_llave(s1).Type = 10 Or Gen_llave(s1).Type = 11 Then
      Gen_llave(s1) = NULO_DATE(txtFields(s1).text)
   Else
    Gen_llave(s1) = Val(txtFields(s1).text)
   End If
   End If
 Next s1
 Gen_llave.Update
 'SALE DEL FORMULARIO
 ACEPTA = "F"
 FRM_STATUS = "0"
 Unload FrmTabla2
 FrmTabla1.Show
End Select

Exit Sub
SALE:
  MsgBox "Error Inesperado . . . # " & Err.Number & "  " & Err.Description
'SALE DEL FORMULARIO
 ACEPTA = "F"
 FRM_STATUS = "0"
 Unload FrmTabla2
 FrmTabla1.Show

End Sub


Private Sub CmdCancela_Click()
'OP_FORM = ""
FRM_STATUS = "0"
Unload FrmTabla2
FrmTabla1.Show

End Sub

Private Sub CmdOtro_Click()
Dim S As Integer
For S = 0 To CAMPOS1 - 1
   If S > 45 Then
    Exit For
   End If
 txtFields(S).text = ""
 Next S
CmdOtro.Enabled = False
CmdAceptar.Enabled = True
If txtFields(0).Enabled Then
 txtFields(0).SetFocus
End If
End Sub

Private Sub Form_Activate()
    FrmTabla2.txtFields(0).SetFocus
End Sub

Private Sub Form_Load()
Dim S As Integer
FrmTabla2.Show
FrmTabla2.SSTab1.Tab = 1
Select Case OP_FORM
Case "X"
  MsgBox "No a Seleccionado NADA ...)Intente Nuevamente... "
  Unload FrmTabla2
Case "A" 'AGREGAR REGISTROS
 CmdOtro.Visible = True
 For S = 0 To CAMPOS1 - 1
   If S > 89 Then
     MsgBox "Solo se mustran 89 campos", 46, Pub_Titulo
     Exit For
   End If
   lblLabels(S).Visible = True
   txtFields(S).Visible = True
   lblLabels(S).Caption = Gen_llave.rdoColumns(S).Name & ":"
'   txtFields(S).MaxLength = Gen_llave.rdoColumns(S).Size
   txtFields(S).TabIndex = S
  Next S
   LblTexto.Caption = "Selección:  Agregar "
Case "D" 'EDITAR CAMPOS
 Gen_llave.AbsolutePosition = Posi_Reg
 For S = 0 To CAMPOS1 - 1
   If S > 89 Then
    MsgBox "Solo se Muestran 89 campos", 46, Pub_Titulo
    Exit For
   End If
   lblLabels(S).Visible = True
   txtFields(S).Visible = True
   lblLabels(S).Caption = Gen_llave.rdoColumns(S).Name & ":"
   txtFields(S).MaxLength = 0 'Gen_llave.rdoColumns(S).Size
   If Not Gen_llave(S) = "" Then
     txtFields(S).text = Gen_llave(S)
   End If
   txtFields(S).TabIndex = S
 Next S
  LblTexto.Caption = "Selección:  Editar "
End Select
FrmTabla2.SSTab1.Tab = 1
LblTexto.Caption = LblTexto.Caption & " >>>   " & wTABLA
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If FrmTabla2.SSTab1.Tab = 1 Then
  If txtFields(0).Visible Then
    txtFields(0).SetFocus
  End If
ElseIf FrmTabla2.SSTab1.Tab = 0 Then
  If txtFields(45).Visible Then
    txtFields(45).SetFocus
  End If
End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
Azul txtFields(Index), txtFields(Index)
End Sub


Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
 CmdCancela.SetFocus
End If
End Sub


