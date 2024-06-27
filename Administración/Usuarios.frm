VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUSUARIOS 
   Caption         =   "Mantenimientos de Usuarios"
   ClientHeight    =   6090
   ClientLeft      =   930
   ClientTop       =   1245
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Icon            =   "Usuarios.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.ListBox lisprecios 
      Height          =   1185
      Left            =   4800
      Style           =   1  'Checkbox
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox listaotro 
      Height          =   1185
      Left            =   6840
      Style           =   1  'Checkbox
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   0
      TabIndex        =   20
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Acceso a Control de Sistema"
      TabPicture(0)   =   "Usuarios.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "F4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "F2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "F3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "F5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ListCias"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Tit"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkcomanda"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Acceso a Facturación"
      TabPicture(1)   =   "Usuarios.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "checont"
      Tab(1).Control(1)=   "FRADEVICE"
      Tab(1).Control(2)=   "fra_serie"
      Tab(1).Control(3)=   "lblfac"
      Tab(1).ControlCount=   4
      Begin VB.CheckBox chkcomanda 
         Caption         =   "Permite Modificar Comandas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   120
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CheckBox checont 
         Caption         =   "Opcinanal "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   -71400
         TabIndex        =   116
         Top             =   360
         Width           =   1815
      End
      Begin VB.ListBox Tit 
         Height          =   1410
         Left            =   3360
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Frame FRADEVICE 
         Caption         =   "Destino de los Documentos (Solo para la Facturación)"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   99
         Top             =   2160
         Width           =   6015
         Begin VB.CheckBox chedevice 
            Caption         =   "Activar Impresion Definida por Usuario."
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   3615
         End
         Begin VB.OptionButton opdevice 
            Caption         =   "Facturas "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   103
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton opdevice 
            Caption         =   "Boletas"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   102
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton opdevice 
            Caption         =   "Guias"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   101
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox cmdimpresoras 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   960
            Width           =   4575
         End
         Begin VB.Label Label3 
            Caption         =   "Impresora (segun en panel de control )"
            Height          =   255
            Left            =   1320
            TabIndex        =   104
            Top             =   720
            Width           =   3015
         End
      End
      Begin VB.ListBox ListCias 
         Height          =   1410
         Left            =   3840
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame F5 
         Caption         =   "Permisos Adicionales en Ventas"
         Height          =   3255
         Left            =   3360
         TabIndex        =   92
         Top             =   1320
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox chkImprime 
            Caption         =   "Imprime Comandas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   2880
            Width           =   2055
         End
         Begin VB.CheckBox chkcuentas 
            Caption         =   "Imprime Precuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CheckBox chkune 
            Caption         =   "Permite Juntar Mesas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1440
            Width           =   2655
         End
         Begin VB.CheckBox chkgrafico 
            Caption         =   "Visualiza Entorno Grafico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkCobrar 
            Caption         =   "Permite Cobrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   2400
            Width           =   1935
         End
         Begin VB.CheckBox chkPassword 
            Caption         =   "Requiere Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   1920
            Width           =   2415
         End
         Begin VB.CheckBox cheventa 
            Caption         =   "Modifica Ventas"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   2160
            TabIndex        =   117
            Top             =   1680
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox chelimite2 
            Caption         =   "Autorizar Creditos Mayor a Permitido"
            Height          =   255
            Left            =   1920
            TabIndex        =   98
            Top             =   840
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.CheckBox chelimite1 
            Caption         =   "Cuentas Pendientes"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton precios 
            Caption         =   "Ver Todos los Precios"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton precios 
            Caption         =   "Variación de Precios"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton precios 
            Caption         =   "Ninguno "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   94
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chestock 
            Caption         =   "Permite Facturar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   2160
            Width           =   3135
         End
      End
      Begin VB.Frame F3 
         Caption         =   "Transacciones de Trabajo :"
         Height          =   2175
         Left            =   3360
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox def 
            Height          =   1635
            Left            =   1680
            Style           =   1  'Checkbox
            TabIndex        =   71
            Top             =   120
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA10"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   23
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   81
            Top             =   1800
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA9"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   22
            Left            =   360
            MaxLength       =   20
            TabIndex        =   80
            Top             =   1800
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA8"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   21
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   79
            Top             =   1410
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA7"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   20
            Left            =   360
            MaxLength       =   20
            TabIndex        =   78
            Top             =   1410
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_COSTRA6"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   19
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   77
            Top             =   1020
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA5"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   18
            Left            =   360
            MaxLength       =   20
            TabIndex        =   76
            Top             =   1020
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA4"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   17
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   75
            Top             =   630
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA3"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   16
            Left            =   360
            MaxLength       =   20
            TabIndex        =   74
            Top             =   630
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA2"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   15
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   73
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_CODTRA1"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   14
            Left            =   360
            MaxLength       =   20
            TabIndex        =   72
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "1.-"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "2.-"
            Height          =   195
            Index           =   15
            Left            =   1680
            TabIndex        =   90
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "3.-"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   89
            Top             =   630
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "4.-"
            Height          =   195
            Index           =   17
            Left            =   1680
            TabIndex        =   88
            Top             =   630
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "5.-"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   87
            Top             =   1020
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "6.-"
            Height          =   195
            Index           =   19
            Left            =   1680
            TabIndex        =   86
            Top             =   1020
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "7.-"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   85
            Top             =   1410
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "8.-"
            Height          =   195
            Index           =   21
            Left            =   1680
            TabIndex        =   84
            Top             =   1410
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "9.-"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   83
            Top             =   1800
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "10.-"
            Height          =   195
            Index           =   23
            Left            =   1680
            TabIndex        =   82
            Top             =   1800
            Width           =   270
         End
      End
      Begin VB.Frame F2 
         Caption         =   "Grupos de Trabajo :"
         Height          =   4215
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ListBox GRUPOS 
            Height          =   2205
            Left            =   1320
            TabIndex        =   39
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO10"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   13
            Left            =   960
            MaxLength       =   2
            TabIndex        =   49
            Top             =   3240
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO9"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   12
            Left            =   960
            MaxLength       =   2
            TabIndex        =   48
            Top             =   2925
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO8"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   11
            Left            =   960
            MaxLength       =   2
            TabIndex        =   47
            Top             =   2595
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO7"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   10
            Left            =   960
            MaxLength       =   2
            TabIndex        =   46
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO6"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   9
            Left            =   960
            MaxLength       =   2
            TabIndex        =   45
            Top             =   1965
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO5"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   8
            Left            =   960
            MaxLength       =   2
            TabIndex        =   44
            Top             =   1635
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO4"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   7
            Left            =   960
            MaxLength       =   2
            TabIndex        =   43
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO3"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   6
            Left            =   960
            MaxLength       =   2
            TabIndex        =   42
            Top             =   1005
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO2"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   5
            Left            =   960
            MaxLength       =   2
            TabIndex        =   41
            Top             =   675
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "USU_GRUPO1"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   4
            Left            =   960
            MaxLength       =   2
            TabIndex        =   40
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO1:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO2:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   68
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO3:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   1005
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO4:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO5:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO6:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   64
            Top             =   1965
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO7:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   63
            Top             =   2280
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO8:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   62
            Top             =   2610
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO9:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   61
            Top             =   2925
            Width           =   720
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "GRUPO10:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   60
            Top             =   3240
            Width           =   810
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   59
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   58
            Top             =   720
            Width           =   765
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   57
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   3
            Left            =   1800
            TabIndex        =   56
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   255
            Index           =   4
            Left            =   1800
            TabIndex        =   55
            Top             =   1680
            Width           =   1245
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   5
            Left            =   1800
            TabIndex        =   54
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   6
            Left            =   1800
            TabIndex        =   53
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   255
            Index           =   7
            Left            =   1800
            TabIndex        =   52
            Top             =   2640
            Width           =   1005
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   195
            Index           =   8
            Left            =   1800
            TabIndex        =   51
            Top             =   3000
            Width           =   885
         End
         Begin VB.Label LBLG 
            AutoSize        =   -1  'True
            Height          =   255
            Index           =   9
            Left            =   1800
            TabIndex        =   50
            Top             =   3240
            Width           =   885
         End
      End
      Begin VB.Frame F4 
         Caption         =   "Menus"
         Height          =   4200
         Left            =   6840
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox TxtTit 
            Height          =   270
            Index           =   7
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   118
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox TxtTit 
            Height          =   285
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   30
            Top             =   285
            Width           =   2295
         End
         Begin VB.TextBox TxtTit 
            Height          =   285
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   29
            Top             =   705
            Width           =   2295
         End
         Begin VB.TextBox TxtTit 
            Height          =   285
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   28
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox TxtTit 
            Height          =   285
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   27
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox TxtTit 
            Height          =   285
            Index           =   5
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   26
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox TxtTit 
            Height          =   285
            Index           =   6
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   25
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox TxtCias 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   3840
            Width           =   2295
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Produccion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   945
            TabIndex        =   119
            Top             =   3150
            Width           =   660
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Mantenimientos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   690
            TabIndex        =   37
            Top             =   120
            Width           =   990
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Controles"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   945
            TabIndex        =   36
            Top             =   960
            Width           =   600
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   930
            TabIndex        =   35
            Top             =   1950
            Width           =   540
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Deliverys"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   960
            TabIndex        =   34
            Top             =   1510
            Width           =   600
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Consultas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   4
            Left            =   960
            TabIndex        =   33
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Procesos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   6
            Left            =   930
            TabIndex        =   32
            Top             =   540
            Width           =   540
         End
         Begin VB.Label Lbltit 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Acceso a Locales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   7
            Left            =   600
            TabIndex        =   31
            Top             =   3650
            Width           =   1050
         End
      End
      Begin VB.Frame fra_serie 
         Caption         =   "Serie Facturación"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   106
         Top             =   360
         Width           =   3255
         Begin VB.TextBox serie_nd 
            Height          =   285
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   110
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox serie_nc 
            Height          =   285
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   109
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox serie_f 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   108
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox serie_b 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   107
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lserie 
            Caption         =   "Serie N. Debito"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   114
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lserie 
            Caption         =   "Serie N. Credito"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   113
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lserie 
            Caption         =   "Serie Facturas"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   112
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lserie 
            Caption         =   "Serie Boletas"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   111
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label lblfac 
         Height          =   255
         Left            =   -74760
         TabIndex        =   115
         Top             =   600
         Width           =   2775
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame F1 
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtFields 
         DataField       =   "USU_KEY"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "USU_NOMBRE"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "USU_CLAVE"
         DataSource      =   "Data1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1080
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "USU_CODCIA"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Usuario :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cia. por Defecto :"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre :"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame F6 
      Height          =   1095
      Left            =   5520
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtprecios 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox usu_fac_imp 
         Caption         =   "Dejar de Imprmir"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtotros 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Ver solo Precios de Venta :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Permisos Especiales"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancelar / Retornar"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Otro"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   6120
      Width           =   975
   End
End
Attribute VB_Name = "frmUSUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim G1 As Integer
Dim IND As Integer
Dim loc_key As Integer

Public Sub MUESTRA_GRUPOS()
Dim cad
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
Public Sub VERIFICA()
Dim i As Integer
For i = 0 To 13
If txtFields(i).text = "" Then
   txtFields(i) = 0
End If
Next i
For i = 14 To 23
If txtFields(i).text = "" Then
   txtFields(i) = " "
End If
Next i

For i = 1 To 5
If TxtTit(i).text = "" Then
   TxtTit(i) = " "
End If
Next i
If txtcias.text = "" Then
  txtcias.text = " "
End If
End Sub


Private Sub checont_Click()
Dim wsc As String
If checont.Value = 0 Then
 wsc = UCase(InputBox("Clave de aceeso: ", "Acceso Limitado", "*********"))
 If wsc <> PUB_CLAVE Then
    checont.Value = 1
    Exit Sub
 End If
End If
End Sub

Private Sub chedevice_Click()
Dim Wflag As Boolean

If chedevice.Value = 1 Then
  Wflag = True
Else
 Wflag = False
End If
opdevice(0).Value = False
opdevice(1).Value = False
opdevice(2).Value = False
opdevice(0).Enabled = Wflag
opdevice(1).Enabled = Wflag
opdevice(2).Enabled = Wflag
cmdimpresoras.ListIndex = -1
cmdimpresoras.Enabled = Wflag
  
  
End Sub

Private Sub cmdAdd_Click()
txtFields(0).text = ""
BORRA_FIELDS 23, txtFields
txtFields(0).SetFocus
cmdAdd.Enabled = False
cmdupdate.Enabled = True
txtcias.text = ""
txtotros.text = ""
precios(0).Value = False
precios(1).Value = False
precios(2).Value = True
chelimite1.Value = 0
chelimite2.Value = 0
chestock.Value = 0
chkcuentas.Value = 0
chkImprime.Value = 0
For fila = 1 To TxtTit.Count
 TxtTit(fila).text = ""
Next fila

End Sub

Private Sub cmdimpresoras_Click()
Dim tem As String * 3
If loc_key = 99 Then Exit Sub
'tem = loc_device
If opdevice(0).Value Then
  opdevice(0).Tag = Trim(Right(cmdimpresoras.text, 4))
ElseIf opdevice(1).Value Then
  opdevice(1).Tag = Trim(Right(cmdimpresoras.text, 4))
ElseIf opdevice(2).Value Then
  opdevice(2).Tag = Trim(Right(cmdimpresoras.text, 4))
End If

End Sub

Private Sub cmdUpdate_Click()
Dim pub_mensaje, estilo, respuesta As String
Dim SPA As String * 30
Dim ww1 As Integer
Screen.MousePointer = 0
If txtFields(0).text = "" Then
   MsgBox "Debe Ingresar Datos del Usuario  ...?  USU_KEY", 48
   txtFields(0).SetFocus
   Exit Sub
 End If
 If txtFields(1).text = "" Then
   MsgBox "Debe Ingresar Datos del Usuario  ...?  USU_NOMBRE", 48
   txtFields(1).SetFocus
   Exit Sub
 End If
 If txtFields(2).text = "" Then
   MsgBox "Debe Ingresar Datos del Usuario  ...?  USU_CLAVE", 48
   txtFields(2).SetFocus
   Exit Sub
 End If
If txtFields(3).text = "" Then
   MsgBox "Debe Ingresar Datos del Usuario  ...?  USU_CODCIA ", 48
   txtFields(3).SetFocus
   Exit Sub
 End If
 SQ_OPER = 1
 PUB_CODCIA = Trim(txtFields(3).text)
 LEER_PAR_LLAVE
 If par_llave.EOF Then
  MsgBox "La Compañia NO existe  ", 48, Pub_Titulo
  Azul txtFields(3), txtFields(3)
  Exit Sub
 End If
 
If chedevice.Value = 1 Then
  If Trim(opdevice(0).Tag) = "" Then
   MsgBox "Definir su Impresora para Facturacion ", 48, Pub_Titulo
   opdevice(0).SetFocus
   Exit Sub
  End If
  If Trim(opdevice(1).Tag) = "" Then
   MsgBox "Definir su Impresora para Facturacion ", 48, Pub_Titulo
   opdevice(1).SetFocus
   Exit Sub
  End If
  If Trim(opdevice(2).Tag) = "" Then
   MsgBox "Definir su Impresora para Facturacion ", 48, Pub_Titulo
   opdevice(2).SetFocus
   Exit Sub
  End If
End If

Screen.MousePointer = 11
'On Error GoTo wlin
If OP_FORM = "A" Then
    usu.Requery
    Do Until usu.EOF
      If Trim(usu!usu_key) = Trim(txtFields(0)) Then
        Screen.MousePointer = 0
        MsgBox "Usuario Ya Exiete ... Intente Nuevamente.  ", 48, Pub_Titulo
        Azul txtFields(0), txtFields(0)
        Exit Sub
      End If
      usu.MoveNext
    Loop
    Gen_llave.AddNew
ElseIf OP_FORM = "D" Then
    Screen.MousePointer = 0
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
ww1 = 0
For ww1 = 0 To 23
    If Gen_llave(ww1).Type = 2 Or Gen_llave(ww1).Type = 3 Or Gen_llave(ww1).Type = 4 Then
     Gen_llave(ww1) = Val(txtFields(ww1).text)
    Else
      Gen_llave(ww1) = txtFields(ww1).text
    End If
Next ww1
For ww1 = 24 To 30 '29
   Gen_llave(ww1) = TxtTit(ww1 - 23).text
   Debug.Print TxtTit(ww1 - 23).text
Next ww1
Gen_llave("USU_MENU7") = TxtTit(7).text
Gen_llave!USU_CIAS = Trim(txtcias.text)
Gen_llave!USU_OTROS = Trim(txtotros.text)
Gen_llave!USU_LISTA_PRE = Trim(txtprecios.text)
Gen_llave!USU_SERIE_B = Trim(Serie_b.text)
Gen_llave!USU_SERIE_F = Trim(serie_f.text)
Gen_llave!USU_SERIE_NC = Trim(serie_nc.text)
Gen_llave!USU_SERIE_ND = Trim(serie_nd.text)
Gen_llave!USU_FLAG_CONT = " "
If checont.Value = 1 Then Gen_llave!USU_FLAG_CONT = "A"



Gen_llave!USU_LIMITE = " "
Gen_llave!usu_stock = " "
If chestock.Value = 1 Then
 Gen_llave!usu_stock = "A"
 Else
 Gen_llave!usu_stock = " "
End If
If chkPassword.Value = 1 Then
 Gen_llave!usu_password = "A"
 Else
 Gen_llave!usu_password = " "
End If

If chkcuentas.Value = 1 Then
 Gen_llave!usu_cuenta = "A"
 Else
 Gen_llave!usu_cuenta = " "
End If

If chkImprime.Value = 1 Then
 Gen_llave!usu_imprime = "A"
 Else
 Gen_llave!usu_imprime = " "
End If

If chelimite1.Value = 1 And chelimite2.Value = 0 Then
 Gen_llave!USU_LIMITE = "A"
End If
If chelimite2.Value = 1 And chelimite1.Value = 0 Then
 Gen_llave!USU_LIMITE = "B"
End If
If chelimite2.Value = 1 And chelimite1.Value = 1 Then
 Gen_llave!USU_LIMITE = "C"
End If
Gen_llave!USU_PRECIO = "C"
If precios(0).Value Then
 Gen_llave!USU_PRECIO = "A"
ElseIf precios(1).Value = True Then
 Gen_llave!USU_PRECIO = "B"
End If
If usu_fac_imp.Value = 1 Then
  Gen_llave!usu_fac_imp = "A"
Else
  Gen_llave!usu_fac_imp = " "
End If
If chedevice.Value = 1 Then
  Gen_llave!USU_DEVICE_FBG = opdevice(0).Tag + opdevice(1).Tag + opdevice(2).Tag
Else
  Gen_llave!USU_DEVICE_FBG = " "
End If

If chkcomanda.Value = 1 Then
  Gen_llave!USU_COMANDA = 1
Else
  Gen_llave!USU_COMANDA = 0
End If

If chkCobrar.Value = 1 Then
  Gen_llave!USU_COBRA = 1
Else
  Gen_llave!USU_COBRA = 0
End If

If chkgrafico.Value = 1 Then
  Gen_llave!USU_FORMACOMANDA = 1
Else
  Gen_llave!USU_FORMACOMANDA = 0
End If

If chkune.Value = 1 Then
  Gen_llave!USU_UNEMESAS = 1
Else
  Gen_llave!USU_UNEMESAS = 0
End If

If OP_FORM = "A" Then
   Gen_llave.Update
   ACEPTA = "F"
   cmdAdd.Enabled = True
   cmdupdate.Enabled = False
   ACEPTA = "F"
ElseIf OP_FORM = "D" Then
  Gen_llave.Update
  Screen.MousePointer = 0
  ACEPTA = "F"
 'SALE DEL FORMULARIO
  FRM_STATUS = "0"
  Unload frmUSUARIOS
  FrmTabla1.Show
End If
Screen.MousePointer = 0

Exit Sub

wlin:
 MsgBox Err.Description
 Screen.MousePointer = 0


End Sub

Private Sub cmdClose_Click()
Screen.MousePointer = 11
FRM_STATUS = "0"
Unload frmUSUARIOS
FrmTabla1.Show 'guiller
Screen.MousePointer = 0
End Sub

Private Sub def_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 27 Then
  def.Visible = False
  txtFields(IND).SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
   txtFields(IND).text = Left(txtFields(IND).text, 4)
   For i = 0 To def.ListCount - 1
     def.ListIndex = i
     If def.Selected(i) Then
         txtFields(IND).text = txtFields(IND).text & "." & Trim(Left(def.text, 2))
     End If
   Next i
   def.Visible = False
   If IND = 23 Then
     txtFields(IND).SetFocus
   Else
    txtFields(IND).SetFocus
   End If
End If

End Sub

Private Sub def_LostFocus()
  def.Visible = False
  txtFields(IND).SetFocus
End Sub

Private Sub Form_Activate()
Dim W1 As Integer
'On Error GoTo sale
frmUSUARIOS.F1.Visible = True
Select Case OP_FORM
Case "X"
     Unload frmUSUARIOS
     MsgBox "Posible conflicto Intente Nuevamente...", 48, Pub_Titulo
  '   FrmTabla1.Show guille
     Exit Sub
Case "A"
  cmdAdd.Visible = True
  cmdAdd.Enabled = False
  txtFields(0).Locked = False
  precios(2).Value = True
  txtFields(0).SetFocus
Case "D" 'EDIATR REGISTROS
  Screen.MousePointer = 11
  cmdAdd.Visible = False
  If LK_CODUSU = "ADMIN" Then
    txtFields(0).Locked = False
  Else
    txtFields(0).Locked = False
  End If
  Gen_llave.AbsolutePosition = Posi_Reg
  W1 = 0
  For W1 = 0 To 23
    txtFields(W1).text = Trim(Gen_llave(W1))
  Next W1
  For W1 = 24 To 30
    TxtTit(W1 - 23).text = Nulo_Valors(Gen_llave(W1))
  Next W1
  precios(0).Value = False
  precios(1).Value = False
  precios(2).Value = False
  If Gen_llave!USU_PRECIO = "A" Then
   precios(0).Value = True
  ElseIf Gen_llave!USU_PRECIO = "B" Then
   precios(1).Value = True
  Else
   precios(2).Value = True
  End If
  chestock.Value = 0
  If Nulo_Valors(Gen_llave!usu_stock) = "A" Then
   chestock.Value = 1
  End If
  
   chkcuentas.Value = 0
  If Nulo_Valors(Gen_llave!usu_cuenta) = "A" Then
   chkcuentas.Value = 1
  End If
  
   chkImprime.Value = 0
  If Nulo_Valors(Gen_llave!usu_imprime) = "A" Then
   chkImprime.Value = 1
  End If
  
  chkPassword.Value = 0
  If Nulo_Valors(Gen_llave!usu_password) = "A" Then
   chkPassword.Value = 1
  End If
  
  chelimite1.Value = 0
  chelimite2.Value = 0
  If Gen_llave!USU_LIMITE = "A" Then
   chelimite1.Value = 1
  End If
  If Gen_llave!USU_LIMITE = "B" Then
   chelimite2.Value = 1
  End If
  If Gen_llave!USU_LIMITE = "C" Then
   chelimite1.Value = 1
   chelimite2.Value = 1
  End If
  
  If Gen_llave!USU_COMANDA = True Then
   chkcomanda.Value = 1
   Else
   chkcomanda.Value = 0
  End If
  
  If Gen_llave!USU_COBRA = True Then
   chkCobrar.Value = 1
   Else
   chkCobrar.Value = 0
  End If
  
   If Gen_llave!USU_FORMACOMANDA = True Then
   chkgrafico.Value = 1
   Else
   chkgrafico.Value = 0
  End If
  
   If Gen_llave!USU_UNEMESAS = True Then
   chkune.Value = 1
   Else
   chkune.Value = 0
  End If
  
  If txtFields(0).text = "ADMIN" Then
    txtFields(0).Enabled = False
  End If
  Serie_b.text = Trim(Nulo_Valors(Gen_llave!USU_SERIE_B))
  serie_f.text = Trim(Nulo_Valors(Gen_llave!USU_SERIE_F))
  serie_nc.text = Trim(Nulo_Valors(Gen_llave!USU_SERIE_NC))
  serie_nd.text = Trim(Nulo_Valors(Gen_llave!USU_SERIE_ND))
  txtcias.text = Nulo_Valors(Gen_llave!USU_CIAS)
  txtotros.text = Nulo_Valors(Gen_llave!USU_OTROS)
  txtprecios.text = Nulo_Valors(Gen_llave!USU_LISTA_PRE)
  If Nulo_Valors(Gen_llave!usu_fac_imp) = "A" Then
    usu_fac_imp.Value = 1
  End If
  opdevice(0).Enabled = False
  opdevice(1).Enabled = False
  opdevice(2).Enabled = False
  cmdimpresoras.ListIndex = -1
  cmdimpresoras.Enabled = False
  If Trim(Nulo_Valors(Gen_llave!USU_DEVICE_FBG)) <> "" Then
    chedevice.Value = 1
    opdevice(0).Tag = Mid(Trim(Nulo_Valors(Gen_llave!USU_DEVICE_FBG)), 1, 1)
    opdevice(1).Tag = Mid(Trim(Nulo_Valors(Gen_llave!USU_DEVICE_FBG)), 2, 1)
    opdevice(2).Tag = Mid(Trim(Nulo_Valors(Gen_llave!USU_DEVICE_FBG)), 3, 1)
    opdevice(0).Value = True
    opdevice_Click 0
  Else
    chedevice.Value = 0
  End If
  checont.Value = 0
  
  If Gen_llave!USU_FLAG_CONT = "A" Then checont.Value = 1
  
  Screen.MousePointer = 0
  txtFields(1).SetFocus
Case Else
  Unload frmUSUARIOS
  MsgBox "Posible conflicto Intente Nuevamente...", 48, Pub_Titulo
'  FrmTabla1.Show guille
  Exit Sub
End Select
MUESTRA_GRUPOS
frmUSUARIOS.F2.Visible = True
frmUSUARIOS.F3.Visible = True
frmUSUARIOS.F4.Visible = True
frmUSUARIOS.F5.Visible = True
frmUSUARIOS.F6.Visible = True
Exit Sub
'sale:
  MsgBox Err.Number & " " & Err.Description & "  " & "LLAMAR A COMPUTO  >>>> FRMUSUARIOS", 48, Pub_Titulo
  Unload frmUSUARIOS
'  FrmTabla1.Show
  Exit Sub

End Sub

Private Sub Form_Load()
Dim P As Printer
fila = 0
For Each P In Printers
    cmdimpresoras.AddItem P.DeviceName & String(80, " ") & fila
    fila = fila + 1
Next P
If LK_FLAG_FACTURACION <> "U" Then
  fra_serie.Visible = False
  lblfac.Visible = True
  If LK_FLAG_FACTURACION = "V" Then
    lblfac.Caption = "Facturación por Vendedores"
  Else
    lblfac.Caption = "Facturación por Compañia"
  End If
End If

'ACEPTA = "'"
'OP_FORM = ""
End Sub

Private Sub GRUPOS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  GRUPOS.Visible = False
  txtFields(IND + 1).SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
  txtFields(IND).text = Val(Right(GRUPOS.text, 2))
  LBLG(IND - 4).Caption = Left(GRUPOS.text, 25)
  GRUPOS.Visible = False
  txtFields(IND + 1).SetFocus
End If
KeyAscii = 0

End Sub

Private Sub lisprecios_KeyPress(KeyAscii As Integer)
Dim i
If KeyAscii = 27 Then
  lisprecios.Visible = False
  txtprecios.SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
   txtprecios.text = ""
   txtprecios.text = Left(txtprecios.text, 4)
   For i = 0 To lisprecios.ListCount - 1
     lisprecios.ListIndex = i
     If lisprecios.Selected(i) Then
         txtprecios.text = txtprecios.text & Val(Trim(Left(lisprecios.text, 2)))
     End If
   Next i
   lisprecios.Visible = False
   If cmdupdate.Enabled And cmdupdate.Visible Then cmdupdate.SetFocus
End If

End Sub

Private Sub lisprecios_LostFocus()
lisprecios.Visible = False
End Sub

Private Sub listaotro_KeyPress(KeyAscii As Integer)
Dim i
If KeyAscii = 27 Then
  listaotro.Visible = False
  txtotros.SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
   txtotros.text = ""
   txtotros.text = Left(txtotros.text, 4)
   For i = 0 To listaotro.ListCount - 1
     listaotro.ListIndex = i
     If listaotro.Selected(i) Then
         txtotros.text = txtotros.text & "." & Trim(Left(listaotro.text, 2))
     End If
   Next i
   listaotro.Visible = False
   If cmdupdate.Enabled And cmdupdate.Visible Then cmdupdate.SetFocus
End If

End Sub

Private Sub listaotro_LostFocus()
listaotro.Visible = False
End Sub

Private Sub ListCias_KeyPress(KeyAscii As Integer)
Dim i
If KeyAscii = 27 Then
  ListCias.Visible = False
  txtcias.SetFocus
  Exit Sub
End If

If KeyAscii = 13 Then
   txtcias.text = ""
   txtcias.text = Left(txtcias.text, 4)
   For i = 0 To ListCias.ListCount - 1
     ListCias.ListIndex = i
     If ListCias.Selected(i) Then
         txtcias.text = txtcias.text & "." & Trim(Left(ListCias.text, 2))
     End If
   Next i
   ListCias.Visible = False
   If txtotros.Enabled And txtotros.Visible Then txtotros.SetFocus
End If


End Sub

Private Sub ListCias_LostFocus()
ListCias.Visible = False
End Sub

Private Sub opdevice_Click(Index As Integer)
Dim WF
Dim wwfg As Integer
WF = opdevice(Index).Tag
wwfg = 0
loc_key = 99
For fila = 0 To cmdimpresoras.ListCount - 1
  cmdimpresoras.ListIndex = fila
  If Trim(Right(cmdimpresoras.text, 4)) = WF Then
    wwfg = 1
    Exit For
  End If
Next fila
If wwfg = 0 Then cmdimpresoras.ListIndex = -1
loc_key = 0

End Sub

Private Sub Tit_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 27 Then
  Tit.Visible = False
  TxtTit(IND).SetFocus
  Exit Sub
End If
If KeyAscii = 13 Then
   TxtTit(IND).text = ""
   TxtTit(IND).text = Left(TxtTit(IND).text, 4)
   For i = 0 To Tit.ListCount - 1
     Tit.ListIndex = i
     If Tit.Selected(i) Then
         TxtTit(IND).text = TxtTit(IND).text & "." & Trim(Left(Tit.text, 2))
     End If
   Next i
   Tit.Visible = False
   If IND <= 5 Then
     ' TxtTit(IND + 1).SetFocus
   Else
      If txtcias.Enabled And txtcias.Visible Then txtcias.SetFocus
   End If
End If

End Sub

Private Sub TxtCias_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    Exit Sub
End If
   LLENA_CIAS 1
   ListCias.Visible = True
   ListCias.SetFocus
End Sub

Private Sub txtFields_Change(Index As Integer)
If txtFields(Index).Index > 3 And txtFields(Index).Index < 14 Then
  Call BUSCA(Index)
End If

End Sub

Private Sub txtFields_DblClick(Index As Integer)
If Index = 2 Then
  MsgBox "Password : " & txtFields(Index).text, 48, Pub_Titulo
End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
If Index < 3 And Index > 14 Then
  Call Azul(txtFields(Index).text, txtFields(Index))
End If

End Sub


Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
Dim a As Integer
If txtFields(Index).Index = 0 Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
If KeyAscii <> 13 Then
   GoTo FIN
End If
If txtFields(Index).Index > 3 And txtFields(Index).Index < 14 Then
   IND = txtFields(Index).Index
   GRUPOS.Visible = True
   GRUPOS.SetFocus
End If
If txtFields(Index).Index > 13 And txtFields(Index).Index < 24 Then
   IND = txtFields(Index).Index
'   LLENA_DEF txtFields(Index).text
   'def_mayor.MoveFirst
   'If def_mayor.EOF Then
   '   MsgBox "NO Tiene Definición de Contabilidad...", 48, Pub_Titulo
   '    Exit Sub
   'End If
   def.Visible = True
   def.SetFocus
End If

FIN:
End Sub



Public Sub BUSCA(VAR1 As Integer)
Dim CRITERIO As String
'CRITERIO = "COD_GRUPO='" & Val(txtFields(4).text) & "'"
If txtFields(VAR1).text = "" Then
  LBLG(VAR1 - 4).Caption = ""
  Exit Sub
End If

For fila = 0 To GRUPOS.ListCount - 1
 GRUPOS.ListIndex = fila
 If Trim(Right(GRUPOS.text, 4)) = txtFields(VAR1).text Then
    LBLG(VAR1 - 4).Caption = Left(GRUPOS.text, 40)
    Exit Sub
  End If
Next fila
Exit Sub

End Sub

Public Sub LLENA_DEF(wcodtra As String)
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
def.Clear
SQ_OPER = 2
PUB_CODTRA = Val(wcodtra)
'LEER_DEF_LLAVE
'If def_mayor.EOF Then
    Exit Sub
'End If
'Do Until def_mayor.EOF
'   DoEvents
'   W1 = def_mayor!DEF_SECUENCIA
'   frmUSUARIOS.def.AddItem W1 & ".-" & def_mayor!DEF_DESCRIPCION
'   def_mayor.MoveNext
'   def.Selected(CUENTA) = False
'   CUENTA = CUENTA + 1
'Loop
cuenta = 0
If Len(Trim(txtFields(IND).text)) = 0 Or Len(Trim(txtFields(IND).text)) = 4 Then
   Exit Sub
End If
WPosV = Len(txtFields(IND).text)
cade = Mid(txtFields(IND).text, 5, WPosV - 4)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
If Mid(cade, i, 1) = "." Then
  a = a + 1
End If
Next i
'a = a - 1

Do Until cuenta = a ' Len(cade) / 2
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   For i = 0 To def.ListCount - 1
     def.ListIndex = i
    If Trim(Left(def.text, 2)) = Trim(WNUM) Then
       def.Selected(i) = True
       Exit For
    End If
   Next i
Loop

End Sub

Public Sub LLENA_TIT(WSTit As Integer)
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
Tit.Clear
PROCESO_TIT WSTit
cuenta = 0
WPosV = Len(TxtTit(IND).text)
cade = Trim(TxtTit(IND).text)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
If Mid(cade, i, 1) = "." Then
  a = a + 1
End If
Next i
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   For i = 0 To Tit.ListCount - 1
     Tit.ListIndex = i
    If Trim(Left(Tit.text, 2)) = Trim(WNUM) Then
       Tit.Selected(i) = True
       Exit For
    End If
   Next i
Loop

End Sub
Public Sub LLENA_CIAS(WSTit As Integer)
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
ListCias.Clear
PROCESO_CIAS
cuenta = 0
WPosV = Len(txtcias.text)
cade = Trim(txtcias.text)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
If Mid(cade, i, 1) = "." Then
  a = a + 1
End If
Next i
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   For i = 0 To ListCias.ListCount - 1
     ListCias.ListIndex = i
    If Trim(Left(ListCias.text, 2)) = Format(CStr(WNUM), "00") Then
       ListCias.Selected(i) = True
       Exit For
    End If
   Next i
Loop

End Sub


Public Sub PROCESO_TIT(wTITULO As Integer)
Dim cuenta As Integer
Dim CUENTA2 As Integer
'On Error GoTo SIGUE
cuenta = 0
Select Case wTITULO
    Case 1
        Do Until cuenta = MDIForm1.SubmenuTit1.Count
        DoEvents
        If MDIForm1.SubmenuTit1.Item(cuenta).Caption <> "-" Then
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.SubmenuTit1.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
        End If
        cuenta = cuenta + 1
        Loop
    Case 2
        Do Until cuenta = MDIForm1.SubmenuTit2.Count
        DoEvents
        If MDIForm1.SubmenuTit2.Item(cuenta).Caption <> "-" Then
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.SubmenuTit2.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
        End If
        cuenta = cuenta + 1
        Loop
    Case 3
        Do Until cuenta = MDIForm1.submenutit3.Count
        DoEvents
        If MDIForm1.submenutit3.Item(cuenta).Caption <> "-" Then
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.submenutit3.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
        End If
        cuenta = cuenta + 1
        Loop
    Case 4
       Do Until cuenta = MDIForm1.SubmenuTit4.Count
        DoEvents
        If MDIForm1.SubmenuTit4.Item(cuenta).Caption <> "-" Then
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.SubmenuTit4.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
        End If
        cuenta = cuenta + 1
        Loop
    Case 5
        Do Until cuenta = MDIForm1.submenutit5.Count
        DoEvents
        If MDIForm1.submenutit5.Item(cuenta).Caption <> "-" Then
          'If LK_CODUSU = "ADMIN" Then
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.submenutit5.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
          'ElseIf cuenta <> 3 Then
          ' frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.submenutit5.Item(cuenta).Caption
          ' Tit.Selected(Tit.ListIndex) = False
          'End If
        End If
        cuenta = cuenta + 1
        Loop
    Case 6
        Do Until cuenta = MDIForm1.SubmenuTit6.Count
        DoEvents
        If MDIForm1.SubmenuTit6.Item(cuenta).Caption <> "-" Then
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.SubmenuTit6.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
        End If
        cuenta = cuenta + 1
        Loop
    Case 7
        Do Until cuenta = MDIForm1.submenutit7.Count
        DoEvents
        Debug.Print MDIForm1.submenutit7.Item(cuenta).Caption
        If MDIForm1.submenutit7.Item(cuenta).Caption <> "-" Then
            
           frmUSUARIOS.Tit.AddItem cuenta & " .-" & MDIForm1.submenutit7.Item(cuenta).Caption
           Tit.Selected(Tit.ListIndex) = False
        End If
        cuenta = cuenta + 1
        Loop
   End Select
Exit Sub
'SIGUE:
'Resume Next
End Sub
Public Sub PROCESO_CIAS()
Dim cuenta As Integer

cuenta = 0
PS_PAR(0) = " "
par.Requery
Do Until par.EOF
     ListCias.AddItem par!PAR_CODCIA & " " & par!PAR_NOMBRE
     cuenta = cuenta + 1
     par.MoveNext
Loop

End Sub

Private Sub txtotros_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    Exit Sub
End If
 LLENA_OTROS 1
 listaotro.Visible = True
 listaotro.SetFocus
End Sub

Private Sub txtprecios_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    Exit Sub
End If
 LLENA_PRECIOS
 lisprecios.Visible = True
 lisprecios.SetFocus

End Sub

Private Sub TxtTit_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then
    Exit Sub
End If
On Error GoTo SALE
   IND = txtFields(Index).Index
   LLENA_TIT Index
   Tit.Visible = True
   Tit.SetFocus
   IND = TxtTit(Index).Index
Exit Sub
SALE:
 MsgBox "Reiniciar Sistema. y Cambiar de usuario . Usuario sin accesos Principales.", 48, Pub_Titulo
End Sub
Public Sub LLENA_OTROS(WSTit As Integer)
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
listaotro.Clear
SQ_OPER = 2
PUB_TIPREG = 111
PUB_CODCIA = "00"
LEER_TAB_LLAVE
listaotro.ToolTipText = "TIPREG = 111"
Do Until tab_mayor.EOF
     listaotro.AddItem Format(tab_mayor!TAB_NUMTAB, "00") & " - " & tab_mayor!TAB_NOMLARGO
     tab_mayor.MoveNext
Loop

cuenta = 0
WPosV = Len(txtotros.text)
cade = Trim(txtotros.text)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
If Mid(cade, i, 1) = "." Then
  a = a + 1
End If
Next i
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   For i = 0 To listaotro.ListCount - 1
     listaotro.ListIndex = i
    If Trim(Left(listaotro.text, 2)) = Format(CStr(WNUM), "00") Then
       listaotro.Selected(i) = True
       Exit For
    End If
   Next i
Loop

End Sub
Public Sub LLENA_PRECIOS()
Dim W1 As String * 2
Dim i, J, wPosF, WPosV, cuenta As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
lisprecios.Clear
SQ_OPER = 2
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
lisprecios.ToolTipText = "TIPREG = 111"
Do Until tab_mayor.EOF
     lisprecios.AddItem Format(tab_mayor!TAB_NUMTAB, "00") & " - " & tab_mayor!TAB_NOMLARGO
     tab_mayor.MoveNext
Loop

cuenta = 0
WPosV = Len(txtprecios.text)
cade = Trim(txtprecios.text)
cuenta = 0
wPosF = 1
a = 0
For i = 1 To Len(cade)
cuenta = Val(Mid(cade, i, 1))
If cuenta <> 0 Then
   For J = 0 To lisprecios.ListCount - 1
     lisprecios.ListIndex = J
     If Val(Left(lisprecios.text, 2)) = cuenta Then
       lisprecios.Selected(J) = True
       Exit For
     End If
   Next J
End If

Next i

End Sub

