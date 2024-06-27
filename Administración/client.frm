VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCLI 
   Caption         =   "Clientes / Proveedores"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   1080
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Frame F14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   51
      Top             =   5880
      Visible         =   0   'False
      Width           =   4125
      Begin MSFlexGridLib.MSFlexGrid ListExiste 
         Height          =   1455
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
      End
      Begin VB.CommandButton CmdEscapa 
         Caption         =   "E&scapar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   56
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdconfirma 
         Caption         =   "Con&firmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   55
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Op 
         Caption         =   "Ignorar la Lista "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   54
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Op 
         Caption         =   "Seleccionar uno de la Lista "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   1800
         Width           =   2535
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   1320
      TabIndex        =   41
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3625
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   57
      Tag             =   "32"
      Top             =   720
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   " "
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " "
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra2"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra1 
         Height          =   3855
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   9255
         Begin VB.CommandButton copia 
            Caption         =   "Copia a Otra Cia."
            Height          =   375
            Left            =   7440
            TabIndex        =   110
            Top             =   1440
            Visible         =   0   'False
            Width           =   1575
         End
         Begin ComctlLib.ProgressBar PB2 
            Height          =   135
            Left            =   7560
            TabIndex        =   106
            Top             =   960
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.TextBox tcuenta2 
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
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdcontab2 
            Caption         =   "Relacionar a Contabilidad"
            Height          =   495
            Left            =   7440
            TabIndex        =   107
            Top             =   840
            Width           =   1815
         End
         Begin ComctlLib.ProgressBar PB 
            Height          =   135
            Left            =   7560
            TabIndex        =   105
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   238
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.CommandButton cmdcontab 
            Caption         =   "Relacionar a Contabilidad"
            Height          =   495
            Left            =   7440
            TabIndex        =   104
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cli_moneda 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   2160
            Visible         =   0   'False
            WhatsThisHelpID =   18
            Width           =   1815
         End
         Begin VB.ComboBox TxtSubZona 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2760
            Width           =   2175
         End
         Begin VB.ComboBox TxtZona 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox txttelefono1 
            DataField       =   "ART_COSPRO"
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
            Height          =   300
            Left            =   5880
            MaxLength       =   12
            TabIndex        =   12
            Text            =   " "
            Top             =   2160
            Width           =   1000
         End
         Begin VB.TextBox txtRUCempresa 
            DataField       =   "ART_UNIDAD"
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
            Height          =   300
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   8
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtRUCesposa 
            DataField       =   "ART_UNIDAD"
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
            Height          =   300
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   6
            Top             =   960
            Width           =   1000
         End
         Begin VB.TextBox Txtnumdir 
            DataField       =   "ART_COSPRO"
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
            Left            =   4920
            MaxLength       =   4
            TabIndex        =   11
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox Txtdireccion 
            DataField       =   "ART_PLANCHA"
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
            Left            =   120
            MaxLength       =   30
            TabIndex        =   9
            Top             =   2160
            Width           =   2775
         End
         Begin VB.TextBox TxtEmpresa 
            DataField       =   "ART_MARGEN"
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
            Left            =   120
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1560
            Width           =   3855
         End
         Begin VB.OptionButton OptNombre 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3600
            TabIndex        =   62
            Top             =   1320
            Width           =   375
         End
         Begin VB.OptionButton OptNombre 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   61
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtesposo 
            DataField       =   "ART_COSTO"
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
            Left            =   120
            MaxLength       =   40
            TabIndex        =   3
            Top             =   360
            Width           =   3855
         End
         Begin VB.OptionButton OptNombre 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   3600
            TabIndex        =   60
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtRUCesposo 
            DataField       =   "ART_UNIDAD"
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
            Height          =   300
            Left            =   4440
            MaxLength       =   15
            TabIndex        =   4
            Top             =   360
            Width           =   1000
         End
         Begin VB.TextBox tcuenta 
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
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   1215
         End
         Begin VB.ListBox ListBloqueos 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   7200
            TabIndex        =   20
            Top             =   2880
            Width           =   1935
         End
         Begin VB.TextBox Txtesposa 
            DataField       =   "ART_IGV"
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
            Left            =   120
            MaxLength       =   40
            TabIndex        =   5
            Top             =   960
            Width           =   3855
         End
         Begin VB.ComboBox cmbgrupo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3360
            Width           =   2295
         End
         Begin VB.ComboBox txtsubgrupo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   3360
            Width           =   2175
         End
         Begin VB.ComboBox txtZonaNew 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2760
            Width           =   2175
         End
         Begin VB.ComboBox TxtLugarCasa 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3000
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtlimite 
            DataField       =   "ART_COSPRO"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4920
            MaxLength       =   12
            TabIndex        =   18
            Top             =   3360
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Cta. Naturaleza:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5880
            TabIndex        =   109
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblfijos 
            AutoSize        =   -1  'True
            Caption         =   "Moneda "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   7200
            TabIndex        =   101
            Top             =   1920
            Visible         =   0   'False
            WhatsThisHelpID =   18
            Width           =   750
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Nombre / Razon Social"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   71
            Tag             =   "1"
            Top             =   120
            Width           =   1995
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "L.E."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   4710
            TabIndex        =   80
            Tag             =   "6"
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "N° Dir."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   4920
            TabIndex        =   79
            Tag             =   "8"
            Top             =   1920
            Width           =   705
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Dirección  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   78
            Tag             =   "7"
            Top             =   1920
            Width           =   1005
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   77
            Tag             =   "10"
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   2520
            TabIndex        =   76
            Tag             =   "11"
            Top             =   2520
            Width           =   810
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Conyuge :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   75
            Tag             =   "5"
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Telefono :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   5880
            TabIndex        =   74
            Tag             =   "9"
            Top             =   1920
            Width           =   885
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "L.E."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   4680
            TabIndex        =   73
            Tag             =   "4"
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Gerente / Representate Legal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Tag             =   "3"
            Top             =   720
            Width           =   2565
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   " -->RUC "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   70
            Tag             =   "2"
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lcuenta 
            Caption         =   "Cta. Activo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5880
            TabIndex        =   69
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label LblDatos 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueos :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   7320
            TabIndex        =   68
            Top             =   2640
            Width           =   915
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Negocio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   67
            Tag             =   "13"
            Top             =   3120
            Width           =   1425
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Clase de Negocio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   13
            Left            =   2520
            TabIndex        =   66
            Tag             =   "14"
            Top             =   3120
            Width           =   1515
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   4920
            TabIndex        =   65
            Tag             =   "12"
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Lugar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   31
            Left            =   3000
            TabIndex        =   64
            Tag             =   "32"
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label lbllimite 
            Caption         =   "Limite de Credito"
            Height          =   255
            Left            =   4920
            TabIndex        =   63
            Top             =   3120
            Width           =   1695
         End
      End
      Begin VB.Frame fra2 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   81
         Top             =   360
         Width           =   9255
         Begin VB.TextBox txtprog 
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
            Left            =   5880
            MaxLength       =   1
            TabIndex        =   37
            Top             =   2280
            Visible         =   0   'False
            WhatsThisHelpID =   19
            Width           =   495
         End
         Begin VB.TextBox txtestado 
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
            Left            =   5880
            MaxLength       =   1
            TabIndex        =   39
            Top             =   3480
            Width           =   495
         End
         Begin VB.CheckBox otrocontrato 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6000
            TabIndex        =   35
            Top             =   840
            Visible         =   0   'False
            WhatsThisHelpID =   7
            Width           =   375
         End
         Begin VB.CheckBox letraotorgado 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   6000
            TabIndex        =   36
            Top             =   1320
            Visible         =   0   'False
            WhatsThisHelpID =   10
            Width           =   435
         End
         Begin VB.TextBox txtDirTrabajo 
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
            Left            =   120
            MaxLength       =   30
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            WhatsThisHelpID =   1
            Width           =   3375
         End
         Begin VB.TextBox txtnumdirtrabajo 
            DataSource      =   "3"
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
            Left            =   6000
            MaxLength       =   4
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            WhatsThisHelpID =   3
            Width           =   615
         End
         Begin VB.ComboBox TxtZonaTrabajo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1080
            Visible         =   0   'False
            WhatsThisHelpID =   5
            Width           =   3375
         End
         Begin VB.ComboBox TxtSubZonaTrabajo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Visible         =   0   'False
            WhatsThisHelpID =   6
            Width           =   1815
         End
         Begin VB.TextBox txtprendas 
            DataField       =   "ART_COLOR"
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
            Left            =   3000
            MaxLength       =   30
            TabIndex        =   34
            Top             =   3480
            Visible         =   0   'False
            WhatsThisHelpID =   17
            Width           =   2580
         End
         Begin VB.TextBox txtpropiedad1 
            DataField       =   "ART_MARGEN"
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
            Left            =   120
            MaxLength       =   30
            TabIndex        =   27
            Top             =   1680
            Visible         =   0   'False
            WhatsThisHelpID =   8
            Width           =   2700
         End
         Begin VB.TextBox txtpropiedad2 
            DataField       =   "ART_MARGEN"
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
            Left            =   120
            MaxLength       =   30
            TabIndex        =   29
            Top             =   2280
            Visible         =   0   'False
            WhatsThisHelpID =   11
            Width           =   2700
         End
         Begin VB.TextBox txtregpublico2 
            DataField       =   "ART_MARGEN"
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
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   30
            Top             =   2280
            Visible         =   0   'False
            WhatsThisHelpID =   12
            Width           =   2580
         End
         Begin VB.TextBox txtautovaluo 
            DataField       =   "ART_MARGEN"
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
            Left            =   120
            MaxLength       =   20
            TabIndex        =   33
            Top             =   3480
            Visible         =   0   'False
            WhatsThisHelpID =   16
            Width           =   2700
         End
         Begin VB.TextBox txtauto1 
            DataField       =   "ART_PLANCHA"
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
            Left            =   120
            MaxLength       =   30
            TabIndex        =   31
            Top             =   2880
            Visible         =   0   'False
            WhatsThisHelpID =   14
            Width           =   2700
         End
         Begin VB.TextBox txtauto2 
            DataField       =   "ART_COSPRO"
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
            Left            =   3000
            MaxLength       =   30
            TabIndex        =   32
            Top             =   2880
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   2580
         End
         Begin VB.TextBox txttelefono2 
            DataSource      =   "4"
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
            Left            =   6840
            MaxLength       =   12
            TabIndex        =   24
            Top             =   480
            Visible         =   0   'False
            WhatsThisHelpID =   4
            Width           =   975
         End
         Begin VB.TextBox txtNucleo 
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
            Left            =   5880
            MaxLength       =   2
            TabIndex        =   38
            Top             =   2880
            Visible         =   0   'False
            WhatsThisHelpID =   13
            Width           =   495
         End
         Begin VB.ComboBox TxtLugarTrab 
            DataSource      =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            WhatsThisHelpID =   2
            Width           =   1815
         End
         Begin VB.TextBox txtDTX 
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
            Left            =   7320
            MaxLength       =   1
            TabIndex        =   40
            Top             =   3480
            Visible         =   0   'False
            WhatsThisHelpID =   19
            Width           =   495
         End
         Begin VB.TextBox txtregpublico1 
            DataField       =   "ART_MARGEN"
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
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   28
            Top             =   1680
            Visible         =   0   'False
            WhatsThisHelpID =   9
            Width           =   2580
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Contrato a Plazo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   27
            Left            =   6480
            TabIndex        =   83
            Tag             =   "28"
            Top             =   960
            Visible         =   0   'False
            WhatsThisHelpID =   7
            Width           =   1515
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Programado :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   34
            Left            =   5880
            TabIndex        =   103
            Tag             =   "35"
            Top             =   2040
            Visible         =   0   'False
            WhatsThisHelpID =   13
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Estado :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   5880
            TabIndex        =   102
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nº. Dir."
            DataSource      =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   15
            Left            =   6000
            TabIndex        =   99
            Tag             =   "16"
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   3
            Width           =   645
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Dirección Trabajo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   98
            Tag             =   "15"
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   1
            Width           =   1530
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   18
            Left            =   3840
            TabIndex        =   97
            Tag             =   "19"
            Top             =   840
            Visible         =   0   'False
            WhatsThisHelpID =   6
            Width           =   810
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   96
            Tag             =   "18"
            Top             =   840
            Visible         =   0   'False
            WhatsThisHelpID =   5
            Width           =   615
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Prendas :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   26
            Left            =   3000
            TabIndex        =   95
            Tag             =   "27"
            Top             =   3240
            Visible         =   0   'False
            WhatsThisHelpID =   17
            Width           =   825
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Prop. (1) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   94
            Tag             =   "20"
            Top             =   1440
            Visible         =   0   'False
            WhatsThisHelpID =   8
            Width           =   870
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Prop. (2) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   93
            Tag             =   "22"
            Top             =   2040
            Visible         =   0   'False
            WhatsThisHelpID =   11
            Width           =   870
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Rg.Pub.(1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   3000
            TabIndex        =   92
            Tag             =   "21"
            Top             =   1440
            Visible         =   0   'False
            WhatsThisHelpID =   9
            Width           =   930
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Rg.Pub.(2)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   22
            Left            =   3000
            TabIndex        =   91
            Tag             =   "23"
            Top             =   2040
            Visible         =   0   'False
            WhatsThisHelpID =   12
            Width           =   930
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Autovaluo :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   90
            Tag             =   "26"
            Top             =   3240
            Visible         =   0   'False
            WhatsThisHelpID =   16
            Width           =   990
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Autos (1) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   89
            Tag             =   "24"
            Top             =   2640
            Visible         =   0   'False
            WhatsThisHelpID =   14
            Width           =   900
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Autos (2) :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   24
            Left            =   3000
            TabIndex        =   88
            Tag             =   "25"
            Top             =   2640
            Visible         =   0   'False
            WhatsThisHelpID =   15
            Width           =   900
         End
         Begin VB.Label lblnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
            DataSource      =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   16
            Left            =   6840
            TabIndex        =   87
            Tag             =   "17"
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   4
            Width           =   885
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Letra Otorgado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   28
            Left            =   6480
            TabIndex        =   86
            Tag             =   "29"
            Top             =   1440
            Visible         =   0   'False
            WhatsThisHelpID =   10
            Width           =   1530
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Opción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   29
            Left            =   5880
            TabIndex        =   85
            Tag             =   "30"
            Top             =   2640
            Visible         =   0   'False
            WhatsThisHelpID =   13
            Width           =   615
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Lugar"
            DataSource      =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   32
            Left            =   3840
            TabIndex        =   84
            Tag             =   "33"
            Top             =   240
            Visible         =   0   'False
            WhatsThisHelpID =   2
            Width           =   615
         End
         Begin VB.Label lblnom 
            AutoSize        =   -1  'True
            Caption         =   "Opcional"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   33
            Left            =   7320
            TabIndex        =   82
            Tag             =   "34"
            Top             =   3240
            Visible         =   0   'False
            WhatsThisHelpID =   19
            Width           =   765
         End
      End
   End
   Begin VB.CommandButton cmdcero 
      Caption         =   "< Crear Codigo con Cero >"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   43
      Top             =   6240
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   48
      Top             =   5400
      Width           =   1600
   End
   Begin VB.TextBox txt_key 
      DataField       =   "ART_KEY"
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
      Left            =   2760
      MaxLength       =   8
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtnombre 
      DataField       =   "ART_NOMBRE"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      MaxLength       =   40
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
   Begin VB.Timer Parpadea 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   6000
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   7800
      TabIndex        =   50
      Top             =   5400
      Width           =   1600
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   5400
      Width           =   1600
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   3960
      TabIndex        =   46
      Top             =   5400
      Width           =   1600
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   2040
      TabIndex        =   44
      Top             =   5400
      Width           =   1600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   2655
      Begin VB.ComboBox CmbCGP 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   12
      X1              =   120
      X2              =   9360
      Y1              =   650
      Y2              =   650
   End
   Begin VB.Label lblnom 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   30
      Left            =   4200
      TabIndex        =   58
      Tag             =   "31"
      Top             =   30
      Width           =   1695
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   45
      Top             =   5160
      Width           =   75
   End
   Begin VB.Label lblvar 
      AutoSize        =   -1  'True
      Caption         =   "Codigo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   47
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_MAX As Integer
Dim llave1 As String
Dim wGARANTES As Integer
Dim CU As Integer
Dim PASA As Integer
Dim UNICO As String
Dim CIA_REF As String * 2
Dim loc_key  As Integer
Dim LOC_TIPREG  As Integer
Dim loc_ultcod As Currency
Dim PSCLILOC_LLAVE As rdoQuery
Dim PSCLILOC_MAYOR As rdoQuery
Dim cliloc_llave As rdoResultset
Dim cliloc_mayor As rdoResultset
Dim PSPAR_CLI As rdoQuery
Dim par_llave_cli As rdoResultset
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim LOC_CANCELA As Integer
Dim LOC_CTA_CLI As String * 12
Dim LOC_DES_CLI As String * 50
Dim LOC_NIVEL As Integer
Dim LOC_CTA_SUP As String
Dim LOC_FLAG_AFEC As String * 1
Dim LOC_ESTADO As String * 1
Dim LOC_TIPO_CTA As Integer
Dim LOC_SIGNO_D As Integer
Dim LOC_SIGNO_H As Integer
Dim LOC_ACT_PAS As Integer

Dim LOC_CTA_CLI2 As String * 12
Dim LOC_DES_CLI2 As String * 50
Dim LOC_NIVEL2 As Integer
Dim LOC_CTA_SUP2 As String
Dim LOC_FLAG_AFEC2 As String * 1
Dim LOC_ESTADO2 As String * 1
Dim LOC_TIPO_CTA2 As Integer
Dim LOC_SIGNO_D2 As Integer
Dim LOC_SIGNO_H2 As Integer
Dim LOC_ACT_PAS2 As Integer



Public Sub LLENA_ZONA(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(15, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENA_GRUPOS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(15, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENA_FILTROS(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem "TODOS " + String(60, "  ") + "T"
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!TAB_NOMLARGO & String(15, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    cont.ListIndex = 0
End Sub


Public Sub LLENA_BLOQ()
   ListBloqueos.Clear
   PUB_TIPREG = 7
   PUB_CODCIA = "00"
   SQ_OPER = 2
   LEER_TAB_LLAVE
   Do Until tab_mayor.EOF
        If cliloc_llave!CLI_TIPO_BLOQ1 = Left(tab_mayor!TAB_NOMCORTO, 1) Then
            ListBloqueos.AddItem tab_mayor!TAB_NOMCORTO
        End If
        If cliloc_llave!CLI_TIPO_BLOQ2 = Left(tab_mayor!TAB_NOMCORTO, 1) Then
            ListBloqueos.AddItem tab_mayor!TAB_NOMCORTO
        End If
        If cliloc_llave!CLI_TIPO_BLOQ3 = Left(tab_mayor!TAB_NOMCORTO, 1) Then
            ListBloqueos.AddItem tab_mayor!TAB_NOMCORTO
        End If
        If cliloc_llave!CLI_TIPO_BLOQ4 = Left(tab_mayor!TAB_NOMCORTO, 1) Then
            ListBloqueos.AddItem tab_mayor!TAB_NOMCORTO
        End If
        tab_mayor.MoveNext
   Loop
End Sub

Public Sub Mayuscula(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If Not car < "a" Or car > "z" Then
  If tecla <> 209 Then
        tecla = 0
        Beep
  End If
End If
End Sub


Public Sub BLOQUEA_TEXT()
    txtnombre.Enabled = False
    txtesposo.Enabled = False
    Txtesposa.Enabled = False
    TxtEmpresa.Enabled = False
    Txtdireccion.Enabled = False
    Txtnumdir.Enabled = False
    TxtZona.Enabled = False
    TxtSubZona.Enabled = False
    txtZonaNew.Enabled = False
    txtDirTrabajo.Enabled = False
    txtnumdirtrabajo.Enabled = False
    frmCLI.TxtZonaTrabajo.Enabled = False
    TxtSubZonaTrabajo.Enabled = False
    txtRUCesposo.Enabled = False
    txtRUCesposa.Enabled = False
    txtRUCempresa.Enabled = False
    frmCLI.txtpropiedad2.Enabled = False
    frmCLI.txtpropiedad1.Enabled = False
    frmCLI.txtregpublico1.Enabled = False
    frmCLI.txtregpublico2.Enabled = False
    frmCLI.txtautovaluo.Enabled = False
    frmCLI.txtauto1.Enabled = False
    frmCLI.txtauto2.Enabled = False
    frmCLI.txtprendas.Enabled = False
    frmCLI.txttelefono1.Enabled = False
    frmCLI.txttelefono2.Enabled = False
    frmCLI.otrocontrato.Enabled = False
    frmCLI.letraotorgado.Enabled = False
    frmCLI.cmbgrupo.Enabled = False
    frmCLI.txtsubgrupo.Enabled = False
    frmCLI.txtNucleo.Enabled = False
    frmCLI.tcuenta.Enabled = False
    frmCLI.OptNombre(0).Enabled = False
    frmCLI.OptNombre(1).Enabled = False
    frmCLI.OptNombre(2).Enabled = False
    frmCLI.txtestado.Enabled = False
    frmCLI.TxtLugarCasa.Enabled = False
    frmCLI.TxtLugarTrab.Enabled = False
    frmCLI.txtlimite.Enabled = False
    cli_moneda.Enabled = False
    txtDTX.Enabled = False
    txtprog.Enabled = False
    cmdcontab.Enabled = False
    cmdcontab2.Enabled = False
    tcuenta2.Enabled = False
End Sub
Public Sub DESBLOQUEA_TEXT()
    txtesposo.Enabled = True
    Txtesposa.Enabled = True
    TxtEmpresa.Enabled = True
    Txtdireccion.Enabled = True
    Txtnumdir.Enabled = True
    TxtZona.Enabled = True
    TxtSubZona.Enabled = True
    txtZonaNew.Enabled = True
    txtDirTrabajo.Enabled = True
    txtnumdirtrabajo.Enabled = True
    frmCLI.TxtZonaTrabajo.Enabled = True
    TxtSubZonaTrabajo.Enabled = True
    txtRUCesposo.Enabled = True
    txtRUCesposa.Enabled = True
    txtRUCempresa.Enabled = True
    frmCLI.txtpropiedad2.Enabled = True
    frmCLI.txtpropiedad1.Enabled = True
    frmCLI.txtregpublico1.Enabled = True
    frmCLI.txtregpublico2.Enabled = True
    frmCLI.txtautovaluo.Enabled = True
    frmCLI.txtauto1.Enabled = True
    frmCLI.txtauto2.Enabled = True
    frmCLI.txtprendas.Enabled = True
    frmCLI.txttelefono1.Enabled = True
    frmCLI.txttelefono2.Enabled = True
    frmCLI.otrocontrato.Enabled = True
    frmCLI.letraotorgado.Enabled = True
    frmCLI.cmbgrupo.Enabled = True
    frmCLI.txtsubgrupo.Enabled = True
    frmCLI.txtNucleo.Enabled = True
    frmCLI.OptNombre(0).Enabled = True
    frmCLI.OptNombre(1).Enabled = True
    frmCLI.OptNombre(2).Enabled = True
    frmCLI.tcuenta.Enabled = True
    frmCLI.txtestado.Enabled = True
    frmCLI.TxtLugarCasa.Enabled = True
    frmCLI.TxtLugarTrab.Enabled = True
    frmCLI.txtlimite.Enabled = True
    cli_moneda.Enabled = True
    txtDTX.Enabled = True
    txtprog.Enabled = True
    cmdcontab.Enabled = True
    cmdcontab2.Enabled = True
    tcuenta2.Enabled = True
End Sub


Private Sub cli_moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.cli_moneda.TabIndex
End If

End Sub

Private Sub CmbCGP_Click()
If llave1 <> "X" Then
  Txt_key.Enabled = False
  If Trim(txtnombre.text) <> "" Then
    LIMPIA_CLI
  End If
  CmbCGP_KeyPress 13
End If
End Sub

Private Sub CmbCGP_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub CmbCGP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CmbCGP.text = "" Then
       CmbCGP.SetFocus
       SendKeys "%{UP}"
       Exit Sub
    End If
'    ALLVISIBLE
    If Left(CmbCGP.text, 1) = "P" Then
      frmCLI.SSTab1.TabCaption(0) = "&Datos Proveedor - Principales"
      frmCLI.SSTab1.TabCaption(1) = "&Datos Proveedor - Opcionales"
       LOC_TIPREG = 310 ' PROVEEDORES
       Screen.MousePointer = 11
       ETIQUETA_CLI
       Screen.MousePointer = 0
       lbllimite.Visible = False
       txtlimite.Visible = False
       lcuenta.Caption = "Cta. Pasivo:"
    Else
      lcuenta.Caption = "Cta. Activo:"
      frmCLI.SSTab1.TabCaption(0) = "&Datos Clientes - Principales"
      frmCLI.SSTab1.TabCaption(1) = "&Datos Clientes - Opcionales"
      LOC_TIPREG = 300 ' CLIENTES
      Screen.MousePointer = 11
      ETIQUETA_CLI
      lbllimite.Visible = False
      txtlimite.Visible = False
      Screen.MousePointer = 0
    End If
    frmCLI.Txt_key.Locked = False
    frmCLI.Txt_key.Enabled = True
    frmCLI.Txt_key.SetFocus
End If
End Sub

Private Sub cmbgrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtsubgrupo.SetFocus
  SendKeys "%{up}"
End If

End Sub

Private Sub cmbgrupo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = cmbgrupo.ListIndex
PUB_TIPREG = Mid(cmbgrupo.ToolTipText, 13, Len(cmbgrupo.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
'DoEvents
LLENA_GRUPOS frmCLI.cmbgrupo, 222
cmbgrupo.SetFocus
SendKeys "%{up}"
fra1.Refresh

End Sub

Private Sub cmdagregar_Click()
On Error GoTo ESCAPA
If Trim(CmbCGP.text) = "" Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(CmdAgregar.Caption, 2) = "&A" And CmdAgregar.Enabled = True Then
    CmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdcero.Enabled = False
    DESBLOQUEA_TEXT
    Txt_key.Locked = True
    LIMPIA_CLI
    If Left(CmbCGP.text, 1) = "C" Then
        frmCLI.OptNombre(0).Value = True
        frmCLI.Txt_key = GENERA_CODI
        If par_llave!PAR_MONEDA_V = "S" Then
          cli_moneda.ListIndex = 0
        ElseIf par_llave!PAR_MONEDA_V = "D" Then
          cli_moneda.ListIndex = 1
        End If
    ElseIf Left(CmbCGP.text, 1) = "P" Then
        frmCLI.OptNombre(0).Value = True
        frmCLI.Txt_key = GENERA_PRO
    End If
    frmCLI.txtesposo.SetFocus
    Txt_key.ToolTipText = ""
    CmbCGP.Enabled = False
    frmCLI.txtestado.text = "A"
    frmCLI.SSTab1.Tab = 0
    PASA = 1
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    'AGREGAMOS EN BLANCO
Else
  If Left(CmbCGP.text, 1) = "C" Then
      If PASA = 1 Then
         If EXISTE_CLI("C", Left(frmCLI.txtesposo.text, 15), Trim(Txt_key.text)) Then
            MENSAJE_CLI " Existen algunos clientes con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
      PASA = 0
      If par_llave!PAR_CONTABILIDAD = "" Then
        GoTo PASACONTAB
      End If
      If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(CmbCGP.text, 1) = "C" Then
            GoTo PASACONTAB
      ElseIf Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(CmbCGP.text, 1) = "P" Then
            GoTo PASACONTAB
      End If
      If Trim(LOC_CTA_CLI) = "" And Trim(LOC_DES_CLI) = "" Then
            If Trim(tcuenta.text) = "" Then
            pub_mensaje = "Desea Relacionarlo a Contabilidad ?"
            Pub_Respuesta = MsgBox(pub_mensaje, vbYesNoCancel + vbQuestion + vbDefaultButton3, Pub_Titulo)
            If Pub_Respuesta = vbYes Then
                cmdcontab_Click
                If LOC_CANCELA = 1 Then
                    MsgBox "Intente Nuevamente colocar Cta. Contable.", 48, Pub_Titulo
                    Exit Sub
                 Else
                    LOC_CANCELA = 0
                End If
            ElseIf Pub_Respuesta = vbCancel Then
               Exit Sub
            End If
         End If
      End If
PASACONTAB:
      If Not CONSIS_CLI Then
          Exit Sub
      End If
     On Error GoTo VERLO_GRABAR
     CN.Execute "Begin Transaction", rdExecDirect
     pub_cadena = "SELECT * FROM CONTROLL"
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     frmCLI.Txt_key = GENERA_CODI
     If wGARANTES = 1 Then
           GRABAR_CLI "G"
     ElseIf wGARANTES = 2 Then
           GRABAR_CLI "H"
     Else
           GRABAR_CLI "C"
     End If
     con_llave.Close
     CN.Execute "Commit Transaction", rdExecDirect
     On Error GoTo 0
     MENSAJE_CLI "Registro,   AGREGADO ... "
     wGARANTES = 0
  ElseIf Left(CmbCGP.text, 1) = "P" Then
      If Trim(cli_moneda.text) = "" Then
        MsgBox "Determine la moneda para el Proveedor", 48, Pub_Titulo
        cli_moneda.SetFocus
        Exit Sub
      End If
      If PASA = 1 Then
         If EXISTE_CLI("P", Left(frmCLI.txtesposo.text, 15), Trim(Txt_key.text)) Then
            MENSAJE_CLI " Existen algunos Proveedor con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
       PASA = 0
       If Not CONSIS_CLI Then
          Exit Sub
       End If
       'SQ_OPER = 1
       'pu_codclie = Val(txt_key.text)
       'pu_cp = "P"
       'pu_codcia = LK_CODCIA
       'LEER_CLILOC_LLAVE
       'If Not cliloc_llave.EOF Then
       '   MsgBox "Proveedor Existe en Compañia ..", 48, Pub_Titulo
       '   Exit Sub
       'End If
       If Trim(LOC_CTA_CLI) = "" And Trim(LOC_DES_CLI) = "" And LK_CODCIA = "03" Then
         If Trim(tcuenta.text) = "" Then
            pub_mensaje = "Antes de Grabar.  Desea Relacionarlo a Contabilidad ?"
            Pub_Respuesta = MsgBox(pub_mensaje, vbYesNoCancel + vbQuestion + vbDefaultButton3, Pub_Titulo)
            If Pub_Respuesta = vbYes Then
                cmdcontab_Click
                If LOC_CANCELA = 1 Then
                    MsgBox "Intente Nuevamente colocar Cta. Contable.", 48, Pub_Titulo
                    Exit Sub
                 Else
                    LOC_CANCELA = 0
                End If
             ElseIf Pub_Respuesta = vbCancel Then
               Exit Sub
            End If
         End If
       End If
       On Error GoTo VERLO_GRABAR
       CN.Execute "Begin Transaction", rdExecDirect
       pub_cadena = "SELECT * FROM CONTROLL"
       Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
       frmCLI.Txt_key = GENERA_PRO
       GRABAR_CLI "P"
       con_llave.Close
       CN.Execute "Commit Transaction", rdExecDirect
       On Error GoTo 0
       MENSAJE_CLI "Proveedor , AGREGADO... "
    End If
    CmdAgregar.Caption = "&Agregar"
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    cmdcero.Enabled = True
    BLOQUEA_TEXT
    Txt_key.Locked = False
    CmbCGP.Enabled = True
    Screen.MousePointer = 0
    frmCLI.SSTab1.Tab = 0
    Txt_key.ToolTipText = ""
    LIMPIA_CLI
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
End If
Exit Sub
    
ESCAPA:
   If Err.Number = 40002 Then
      Screen.MousePointer = 0
      MsgBox "El Codigo generado ya existe " & Chr(13) & "Se procede a generar el siguiente codigo y a continuación " & Chr(13) & "Intente Grabar Nuevamente...", 48, Pub_Titulo
      frmCLI.Txt_key = GENERA_CODI
      Resume Next
      Exit Sub
   Else
      Screen.MousePointer = 0
      MsgBox Err.Number & "  " & Err.Description & "   ...  LLAMAR A COMPUTO"
      cmdcancelar_Click
      
   End If
   Exit Sub
VERLO_GRABAR:
'    If con_llave Is Nothing Then
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
'    End If
    cmdcancelar_Click
FIN:
End Sub

Private Sub cmdagregar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   If frmCLI.Txt_key.Visible Then
      frmCLI.Txt_key.SetFocus
   End If
End If

End Sub

Private Sub cmdcancelar_Click()
If Txt_key.Visible = False Then
  Exit Sub
End If
If Left(CmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_CLI
    cmdCancelar.Enabled = True
    Txt_key.Locked = False
    MENSAJE_CLI "Proceso Cancelado... !!!    "
    Txt_key.Enabled = True
    Txt_key.SetFocus
    frmCLI.SSTab1.Tab = 0
    Screen.MousePointer = 0
    PASA = 0
    cmdcontab.Enabled = False
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
    Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        If Left(CmbCGP.text, 1) = "C" Then
           LLENA_CLI 1, "C"
        Else
           LLENA_CLI 1, "P"
        End If
        Txt_key.Locked = True
     Else
        CmdAgregar.Caption = "&Agregar"
        cmdcontab.Enabled = False
        LIMPIA_CLI
        Txt_key.Locked = False
        Txt_key.SetFocus
     End If
     CmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     cmdModificar.Enabled = True
     cmdcero.Enabled = True
     Txt_key.ToolTipText = ""
     wGARANTES = 0
     BLOQUEA_TEXT
     MENSAJE_CLI "Proceso Cancelado... !!!    "
     CmbCGP.Enabled = True
     frmCLI.SSTab1.Tab = 0
     Screen.MousePointer = 0
     LOC_CTA_CLI = ""
     LOC_CTA_CLI2 = ""
     PASA = 0
End Sub

Private Sub cmdCancelar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdcero_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdcerrar_Click()
Dim iFormCount As Integer
Dim WCODI As String
cmdcancelar_Click
frmCLI.Hide

End Sub

Private Sub cmdCerrar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.Txt_key.SetFocus
End If
End Sub

Private Sub cmdconfirma_Click()
  If Op(0).Value And Left(frmCLI.CmbCGP, 1) = "C" Then
     frmCLI.Txt_key.text = ListExiste.TextMatrix(ListExiste.Row, 1)
     PASA = 1
     frmCLI.F14.Visible = False
     cmdagregar_Click
     Exit Sub
  End If
  If Op(0).Value And Left(frmCLI.CmbCGP, 1) = "P" Then
    frmCLI.txtnombre.text = ListExiste.TextMatrix(ListExiste.Row, 2)
    frmCLI.Txt_key.text = ListExiste.TextMatrix(ListExiste.Row, 1)
     PASA = 1
     frmCLI.F14.Visible = False
     If Left(CmdAgregar.Caption, 2) = "&G" And CmdAgregar.Enabled = True Then cmdagregar_Click
     If Left(cmdModificar.Caption, 2) = "&G" And cmdModificar.Enabled = True Then CmdModificar_Click
     Exit Sub
  End If
  If Op(1).Value Then
     PASA = 0
     frmCLI.F14.Visible = False
     If Left(CmdAgregar.Caption, 2) = "&G" And CmdAgregar.Enabled = True Then cmdagregar_Click
     If Left(cmdModificar.Caption, 2) = "&G" And cmdModificar.Enabled = True Then CmdModificar_Click
     Exit Sub
  End If
  MsgBox "Seleccione una de las dos Opciones ..", 48, Pub_Titulo
End Sub

Private Sub cmdcontab_Click()
If par_llave!PAR_CONTABILIDAD <> "A" Then
  Exit Sub
End If
If Left(CmbCGP.text, 1) = "C" Then
  If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(cmdcontab.Caption, 2) = "&Q" Then
      tcuenta.text = ""
      cmdcontab.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
ElseIf Left(CmbCGP.text, 1) = "P" Then
  If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(cmdcontab.Caption, 2) = "&Q" Then
      tcuenta.text = ""
      cmdcontab.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If
If Left(cmdcontab.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & tcuenta.text & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
    SQ_OPER = 1
    PUB_CUENTA = Trim(tcuenta.text)
    LEER_COM_LLAVE
    If com_llave.EOF Then
      tcuenta.text = ""
    Else
      com_llave.Delete
      tcuenta.text = ""
      CmdModificar_Click
    End If
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    Exit Sub
End If
LOC_CANCELA = 0
If txtesposo.text = "" Then
 MsgBox "Ingrese Descripción del cliente..", 48, Pub_Titulo
 Azul txtesposo, txtesposo
 Exit Sub
End If
If Left(CmbCGP.text, 1) = "C" Then
    LK_TABLA = "CLIENTE"
    archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND ( COM_CUENTA >= '" & 12 & "' AND COM_CUENTA < '" & 13 & "' OR COM_CUENTA >= '" & 16 & "' AND COM_CUENTA < '" & 17 & "' OR COM_CUENTA >= '" & 14 & "' AND COM_CUENTA < '" & 15 & "' ) ORDER BY COM_CUENTA"
Else
   LK_TABLA = "PROVEEDOR"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 42 & "' AND COM_CUENTA < '" & 43 & "' OR COM_CUENTA >= '" & 46 & "' AND COM_CUENTA < '" & 47 & "' )  ORDER BY COM_CUENTA"
End If
LOC_CTA_CLI = ""
LOC_DES_CLI = ""
PB.Visible = True
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
PB.Visible = False
frmBuscacta.Show 1
LOC_CTA_CLI = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI = Trim(frmBuscacta.tnombre.text)
LOC_NIVEL = Val(frmBuscacta.txtdatos(0).text)
LOC_CTA_SUP = Trim(frmBuscacta.txtdatos(1).text)
LOC_FLAG_AFEC = Trim(frmBuscacta.txtdatos(2).text)
LOC_ESTADO = Trim(frmBuscacta.txtdatos(3).text)
LOC_TIPO_CTA = Val(frmBuscacta.txtdatos(4).text)
LOC_SIGNO_D = Val(frmBuscacta.txtdatos(5).text)
LOC_SIGNO_H = Val(frmBuscacta.txtdatos(6).text)
LOC_ACT_PAS = Val(frmBuscacta.txtdatos(7).text)
tcuenta = Trim(LOC_CTA_CLI)
If Trim(LOC_DES_CLI) = "" And Trim(LOC_DES_CLI) = "" Then
 LOC_CANCELA = 1
Else
 LOC_CANCELA = 0
End If
Unload frmBuscacta
If LOC_CANCELA = 1 Then
 Exit Sub
End If
If Left(CmbCGP.text, 1) = "C" Then
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" Then
     Exit Sub
   End If
ElseIf Left(CmbCGP.text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
End If

If Left(cmdModificar.Caption, 2) = "&G" Then
   CmdModificar_Click
End If
End Sub

Private Sub cmdcontab2_Click()
If par_llave!PAR_CONTABILIDAD <> "A" Then
  Exit Sub
End If
If Left(CmbCGP.text, 1) = "P" Then
  If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(cmdcontab2.Caption, 2) = "&Q" Then
      tcuenta2.text = ""
      cmdcontab2.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If
If Left(CmbCGP.text, 1) = "C" Then
  If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(cmdcontab2.Caption, 2) = "&Q" Then
      tcuenta2.text = ""
      cmdcontab2.Caption = "Relacionar a Con&tabilidad"
      Exit Sub
  End If
End If

If Left(cmdcontab2.Caption, 2) = "&Q" Then
    pub_mensaje = "Confirmar la eliminación de la Cuenta : " & tcuenta2.text & " , Continuar ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If
     tcuenta2.text = ""
     CmdModificar_Click
     cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    Exit Sub
End If
LOC_CANCELA = 0
If txtesposo.text = "" Then
 MsgBox "Ingrese Descripción del cliente..", 48, Pub_Titulo
 Azul txtesposo, txtesposo
 Exit Sub
End If
If Left(CmbCGP.text, 1) = "C" Then
   LK_TABLA = "CLIENTES2"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 70 & "' AND COM_CUENTA < '" & 71 & "' OR COM_CUENTA >= '" & 75 & "' AND COM_CUENTA < '" & 78 & "') ORDER BY COM_CUENTA"
End If
If Left(CmbCGP.text, 1) = "P" Then
   LK_TABLA = "PROVEEDOR2"
   archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= '" & 60 & "' AND COM_CUENTA < '" & 61 & "' OR COM_CUENTA >= '" & 62 & "' AND COM_CUENTA < '" & 68 & "' OR COM_CUENTA >= '" & 33 & "' AND COM_CUENTA < '" & 39 & "' ) ORDER BY COM_CUENTA"
End If
LOC_CTA_CLI2 = ""
LOC_DES_CLI2 = ""
PB2.Visible = True
DoEvents
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
PB.Visible = False
frmBuscacta.Show 1
LOC_CTA_CLI2 = Trim(frmBuscacta.tcuenta)
LOC_DES_CLI2 = Trim(frmBuscacta.tnombre.text)
LOC_NIVEL2 = Val(frmBuscacta.txtdatos(0).text)
LOC_CTA_SUP2 = Trim(frmBuscacta.txtdatos(1).text)
LOC_FLAG_AFEC2 = Trim(frmBuscacta.txtdatos(2).text)
LOC_ESTADO2 = Trim(frmBuscacta.txtdatos(3).text)
LOC_TIPO_CTA2 = Val(frmBuscacta.txtdatos(4).text)
LOC_SIGNO_D2 = Val(frmBuscacta.txtdatos(5).text)
LOC_SIGNO_H2 = Val(frmBuscacta.txtdatos(6).text)
LOC_ACT_PAS2 = Val(frmBuscacta.txtdatos(7).text)
tcuenta2 = Trim(LOC_CTA_CLI2)
If Trim(LOC_DES_CLI2) = "" And Trim(LOC_DES_CLI2) = "" Then
 LOC_CANCELA = 1
Else
 LOC_CANCELA = 0
End If
Unload frmBuscacta
If LOC_CANCELA = 1 Then
 Exit Sub
End If
If Left(CmbCGP.text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
End If

If Left(cmdModificar.Caption, 2) = "&G" Then
   CmdModificar_Click
End If

End Sub

Private Sub cmdEliminar_Click()
Dim wcias As String
On Error GoTo SALE
If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
  Dim PS_REP01 As rdoQuery
  Dim llave_rep01 As rdoResultset
  Screen.MousePointer = 11
  LblMensaje.Visible = True
  LblMensaje.Caption = "Verificando Data.  un Momento..."
  pub_cadena = "SELECT FAR_CODCLIE FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODCLIE = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = cliloc_llave!CLI_CODCLIE
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     LblMensaje.Visible = False
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  CLIENTE  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  Screen.MousePointer = 0
  LblMensaje.Caption = ""
  If Trim(Nulo_Valors(gen!gen_cli_cias)) <> "" Then
    wcias = Trim(gen!gen_cli_cias)
    MsgBox "O J O ...  Al Eliminar este Cliente tambien debe hacerlo con las demas Compañias relacionadas : " & wcias, 48, Pub_Titulo
  End If
  If Trim(tcuenta.text) <> "" Then
    pub_mensaje = " ¿Desea Eliminar el Registro, y su Relacion a Contabilidad .. ?"
  Else
    pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  End If
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    cliloc_llave.Delete
    frmCLI.Txt_key.text = ""
    frmCLI.Txt_key.Locked = False
    If Trim(tcuenta.text) <> "" Then
     SQ_OPER = 1
     PUB_CUENTA = Trim(tcuenta.text)
     LEER_COM_LLAVE
     If com_llave.EOF Then
         tcuenta.text = ""
     Else
         com_llave.Delete
         tcuenta.text = ""
     End If
    End If
    If Trim(tcuenta2.text) <> "" Then
     SQ_OPER = 1
     PUB_CUENTA = Trim(tcuenta2.text)
     LEER_COM_LLAVE
     If com_llave.EOF Then
         tcuenta2.text = ""
     Else
         com_llave.Delete
         tcuenta2.text = ""
     End If
    End If
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    LIMPIA_CLI
    MENSAJE_CLI "Registro   ELIMINADO ... "
    Screen.MousePointer = 0
  End If
  Screen.MousePointer = 0
Exit Sub
SALE:
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    cmdcancelar_Click
    Screen.MousePointer = 0

End Sub

Private Sub cmdEliminar_GotFocus()
If ListView1.Visible Then
frmCLI.Txt_key.text = ""
frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub cmdEliminar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.Txt_key.SetFocus
End If

End Sub

Private Sub CmdEscapa_Click()
  frmCLI.F14.Visible = False
'  PASA = 0
  If frmCLI.txtesposo.Enabled Then
      frmCLI.txtesposo.SetFocus
  End If
End Sub

Private Sub CmdModificar_Click()
If Len(Txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_CLI "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdEliminar.Enabled = False
    CmdAgregar.Enabled = False
    cmdcero.Enabled = False
    cmdCancelar.Enabled = True
    CmbCGP.Enabled = False
    DESBLOQUEA_TEXT
    Txt_key.Locked = True
    frmCLI.txtesposo.SetFocus
    PASA = 1
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""

 Else
   If Left(CmbCGP.text, 1) = "C" Then
      If PASA = 1 Then
         If EXISTE_CLI("C", Left(frmCLI.txtesposo.text, 15), Trim(Txt_key.text)) Then
            MENSAJE_CLI " Existen algunos clientes con estos NOMBRES .."
            frmCLI.ListExiste.SetFocus
            Exit Sub
         End If
      End If
      PASA = 0
   ElseIf Left(CmbCGP.text, 1) = "P" Then
     If PASA = 1 Then
      If EXISTE_CLI("P", Left(frmCLI.txtesposo.text, 15), Trim(Txt_key.text)) Then
         MENSAJE_CLI " Existen algunos Proveedor con estos NOMBRES .."
         frmCLI.ListExiste.SetFocus
         Exit Sub
      End If
    End If
    PASA = 0
   End If
   If Not CONSIS_CLI Then
         '  "NO SE PUEDE.."
      Exit Sub
   End If
   If par_llave!PAR_CONTABILIDAD = "" Then
      GoTo PASACONTAB
   End If
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" And Left(CmbCGP.text, 1) = "C" Then
      GoTo PASACONTAB
   ElseIf Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" And Left(CmbCGP.text, 1) = "P" Then
      GoTo PASACONTAB
   End If
   If Left(cmdcontab.Caption, 2) <> "&Q" Or Left(cmdcontab2.Caption, 2) <> "&Q" Then
      If Trim(tcuenta.text) <> "" Or Trim(tcuenta2.text) <> "" Then
         GRABA_CONTAB LK_CODCIA
      End If
      
   End If
PASACONTAB:
    Screen.MousePointer = 11
    GRABAR_CLI "C"
    MENSAJE_CLI "Registro , MODIFICADO... "
    cmdModificar.Caption = "&Modificar"
    frmCLI.SSTab1.Tab = 0
    Screen.MousePointer = 0
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    CmdAgregar.Enabled = True
    cmdcero.Enabled = True
    BLOQUEA_TEXT
    Txt_key.Locked = True
    CmbCGP.Enabled = True
    cmdCancelar.SetFocus
    Screen.MousePointer = 0
    LOC_CTA_CLI = ""
    LOC_CTA_CLI2 = ""
  
End If
End Sub

Private Sub cmdModificar_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub CmdCero_Click()
  pub_mensaje = " ¿Desea Generar el " & Mid(frmCLI.CmbCGP.text, 3, Len(frmCLI.CmbCGP.text)) & " con codigo 0 ..?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then   ' El usuario eligió
    MENSAJE_CLI "Proceso Cancelado ..."
    Exit Sub
  End If
    SQ_OPER = 1
    pu_cp = " "
    pu_codclie = 0
    pu_codcia = LK_CODCIA
    LEER_CLILOC_LLAVE
    If Not cliloc_llave.EOF Then
      MsgBox "codigo Cero YA existe en CIA ...", 48, Pub_Titulo
      Exit Sub
    End If
    cliloc_llave.AddNew
    cliloc_llave!CLI_CP = " "
    cliloc_llave!CLI_CODCLIE = 0
    cliloc_llave!CLI_CODCIA = LK_CODCIA
    cliloc_llave!CLI_NOMBRE_ESPOSO = " Ninguno"
    cliloc_llave!CLI_NOMBRE_ESPOSA = " Ninguno"
    cliloc_llave!CLI_NOMBRE_EMPRESA = " Ninguno"
    cliloc_llave!CLI_123 = 1
    cliloc_llave!cli_nombre = " Ninguno"
    cliloc_llave.Update
    MENSAJE_CLI "Codigo   Generado ..."
End Sub

Private Sub copia_Click()
Dim valor
   If Val(frmCLI.Txt_key.text) <= 0 Then
      MsgBox " Consulte  y  despues Copiar.."
      Exit Sub
   End If
   
    valor = InputBox("La Compañia a donde copiar los datos : ", "COMPAÑIA", "03")
    If valor = "" Then Exit Sub
    If Trim(valor) = LK_CODCIA Then
       MsgBox "No Procede .. "
       Exit Sub
    End If
    If Len(Trim(valor)) <> 2 Then
       MsgBox "de dos digitos  "
       Exit Sub
    End If
    
    cliloc_llave.AddNew
    cliloc_llave!CLI_CP = Left(CmbCGP.text, 1)
    cliloc_llave!CLI_CODCLIE = Val(frmCLI.Txt_key.text)
    cliloc_llave!cli_SALDO = 0
    cliloc_llave!CLI_DET_TOT = "D"
    cliloc_llave!cli_moneda = "S"
    cliloc_llave!CLI_CODCIA = valor
    cliloc_llave!CLI_NOMBRE_ESPOSO = txtesposo.text
    cliloc_llave!CLI_NOMBRE_ESPOSA = Txtesposa.text
    cliloc_llave!CLI_NOMBRE_EMPRESA = TxtEmpresa.text
    ASIGNA_123
    cliloc_llave!cli_nombre = frmCLI.txtnombre.text
    cliloc_llave!CLI_CASA_DIREC = Txtdireccion.text
    cliloc_llave!CLI_CASA_NUM = Val(Txtnumdir.text)
    cliloc_llave!CLI_CASA_ZONA = Val(Right(TxtZona.text, 4))
    cliloc_llave!CLI_LUGAR_CASA = Val(Right(TxtLugarCasa.text, 4))
    cliloc_llave!CLI_LUGAR_TRAB = Val(Right(TxtLugarTrab.text, 4))
    cliloc_llave!CLI_CASA_SUBZONA = Val(Right(TxtSubZona.text, 4))
    cliloc_llave!CLI_ZONA_NEW = Val(Right(txtZonaNew.text, 4))
    cliloc_llave!CLI_TRAB_DIREC = txtDirTrabajo.text
    cliloc_llave!CLI_TRAB_NUM = Nulo_Valor0(txtnumdirtrabajo.text)
    cliloc_llave!cli_TRAB_ZONA = Val(Right(frmCLI.TxtZonaTrabajo.text, 4))
    cliloc_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.text, 4))
    cliloc_llave!cli_RUC_ESPOSO = txtRUCesposo.text
    cliloc_llave!CLI_RUC_ESPOSA = txtRUCesposa.text
    cliloc_llave!CLI_RUC_EMPRESA = txtRUCempresa.text
    cliloc_llave!CLI_CASA1 = frmCLI.txtpropiedad1.text
    cliloc_llave!CLI_CASA2 = frmCLI.txtpropiedad2.text
    cliloc_llave!CLI_REGPUB1 = frmCLI.txtregpublico1.text
    cliloc_llave!CLI_REGPUB2 = frmCLI.txtregpublico2.text
    cliloc_llave!CLI_AUTOAVALUO = frmCLI.txtautovaluo.text
    cliloc_llave!CLI_AUTO1 = frmCLI.txtauto1.text
    cliloc_llave!CLI_AUTO2 = frmCLI.txtauto2.text
    cliloc_llave!CLI_PRENDA = frmCLI.txtprendas.text
    cliloc_llave!CLI_TELEF1 = frmCLI.txttelefono1.text
    cliloc_llave!CLI_TELEF2 = frmCLI.txttelefono2.text
    cliloc_llave!CLI_OTRO_CONTR = frmCLI.otrocontrato.Value
    cliloc_llave!CLI_LETRA = frmCLI.letraotorgado.Value
    cliloc_llave!CLI_GRUPO = Val(Right(frmCLI.cmbgrupo.text, 4))
    cliloc_llave!CLI_SUBGRUPO = Val(Right(frmCLI.txtsubgrupo.text, 4))
    cliloc_llave!CLI_CORRELA = 0
    cliloc_llave!CLI_nucleo = frmCLI.txtNucleo.text
    cliloc_llave!CLI_estado = frmCLI.txtestado.text
    cliloc_llave!CLI_programado = Nulo_Valors(txtprog.text)
    '  <<< Actualiza La Cta. solo de la Cia Actual >>>
    cliloc_llave!CLI_CUENTA_CONTAB = Trim(frmCLI.tcuenta.text)
    cliloc_llave!CLI_CUENTA_CONTAB2 = Trim(frmCLI.tcuenta2.text)
    cliloc_llave!cli_moneda = Left(cli_moneda.text, 1)
    If txtDTX.text = "" Then
    txtDTX.text = " "
    End If
    cliloc_llave!CLI_DET_TOT = txtDTX.text
    cliloc_llave!cli_limcre = Val(txtlimite.text)
cliloc_llave.Update
MsgBox "Proceso Copiado .... ", 48, Pub_Titulo
Unload frmCLI
End Sub

Private Sub Form_Activate()
'frmCLI.Refresh
End Sub

Private Sub Form_Load()
Dim i As Integer
End Sub

Public Sub ALLINVISIBLE()
    frmCLI.lcuenta.Visible = False
    Txt_key.Visible = False
    txtnombre.Visible = False
    txtesposo.Visible = False
    Txtesposa.Visible = False
    TxtEmpresa.Visible = False
    Txtdireccion.Visible = False
    Txtnumdir.Visible = False
    TxtZona.Visible = False
    TxtSubZona.Visible = False
    txtZonaNew.Visible = False
    txtDirTrabajo.Visible = False
    txtnumdirtrabajo.Visible = False
    frmCLI.TxtZonaTrabajo.Visible = False
    TxtSubZonaTrabajo.Visible = False
    txtRUCesposo.Visible = False
    txtRUCesposa.Visible = False
    txtRUCempresa.Visible = False
    frmCLI.txtpropiedad2.Visible = False
    frmCLI.txtpropiedad1.Visible = False
    frmCLI.txtregpublico1.Visible = False
    frmCLI.txtregpublico2.Visible = False
    frmCLI.txtautovaluo.Visible = False
    frmCLI.txtauto1.Visible = False
    frmCLI.txtauto2.Visible = False
    frmCLI.txtprendas.Visible = False
    frmCLI.txttelefono1.Visible = False
    frmCLI.txttelefono2.Visible = False
    frmCLI.otrocontrato.Visible = False
    frmCLI.letraotorgado.Visible = False
    frmCLI.ListBloqueos.Visible = False
    frmCLI.OptNombre(0).Visible = False
    frmCLI.OptNombre(1).Visible = False
    frmCLI.OptNombre(2).Visible = False
    frmCLI.cmbgrupo.Visible = False
    frmCLI.txtsubgrupo.Visible = False
    frmCLI.txtNucleo.Visible = False
    frmCLI.tcuenta.Visible = False
    frmCLI.txtestado.Visible = False
    frmCLI.TxtLugarCasa.Visible = False
    frmCLI.TxtLugarTrab.Visible = False
End Sub
Public Sub ALLVISIBLE()
    frmCLI.lcuenta.Visible = True
    Txt_key.Visible = True
    txtnombre.Visible = True
    txtesposo.Visible = True
    Txtesposa.Visible = True
    TxtEmpresa.Visible = True
    Txtdireccion.Visible = True
    Txtnumdir.Visible = True
    TxtZona.Visible = True
    TxtSubZona.Visible = True
    txtZonaNew.Visible = True
    txtDirTrabajo.Visible = True
    txtnumdirtrabajo.Visible = True
    frmCLI.TxtZonaTrabajo.Visible = True
    TxtSubZonaTrabajo.Visible = True
    txtRUCesposo.Visible = True
    txtRUCesposa.Visible = True
    txtRUCempresa.Visible = True
    frmCLI.txtpropiedad2.Visible = True
    frmCLI.txtpropiedad1.Visible = True
    frmCLI.txtregpublico1.Visible = True
    frmCLI.txtregpublico2.Visible = True
    frmCLI.txtautovaluo.Visible = True
    frmCLI.txtauto1.Visible = True
    frmCLI.txtauto2.Visible = True
    frmCLI.txtprendas.Visible = True
    frmCLI.txttelefono1.Visible = True
    frmCLI.txttelefono2.Visible = True
    frmCLI.otrocontrato.Visible = True
    frmCLI.letraotorgado.Visible = True
    frmCLI.ListBloqueos.Visible = True
    frmCLI.OptNombre(0).Visible = True
    frmCLI.OptNombre(1).Visible = True
    frmCLI.OptNombre(2).Visible = True
    frmCLI.cmbgrupo.Visible = True
    frmCLI.txtsubgrupo.Visible = True
    frmCLI.txtNucleo.Visible = True
    frmCLI.tcuenta.Visible = True
    frmCLI.txtestado.Visible = True
    frmCLI.TxtLugarCasa.Visible = True
    frmCLI.TxtLugarTrab.Visible = True
End Sub

Public Sub LLENA_123()
  If cliloc_llave!CLI_123 = 1 Then
       frmCLI.OptNombre(0).Value = True
  ElseIf cliloc_llave!CLI_123 = 2 Then
       frmCLI.OptNombre(1).Value = True
  ElseIf cliloc_llave!CLI_123 = 3 Then
       frmCLI.OptNombre(2).Value = True
       Exit Sub
  End If

End Sub
Public Sub ASIGNA_123()
  If frmCLI.OptNombre(0).Value Then
     frmCLI.txtnombre.text = Nulo_Valors(frmCLI.txtesposo.text)
     cliloc_llave!CLI_123 = 1
  ElseIf frmCLI.OptNombre(1).Value Then
     frmCLI.txtnombre.text = Nulo_Valors(frmCLI.Txtesposa.text)
     cliloc_llave!CLI_123 = 2
  ElseIf frmCLI.OptNombre(2).Value Then
     frmCLI.txtnombre.text = Nulo_Valors(frmCLI.TxtEmpresa.text)
     cliloc_llave!CLI_123 = 3
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'cli_mayor2.Close
 fila = 0
 pub_cadena = ""
End Sub

Private Sub lblnom_DblClick(Index As Integer)
If Trim(LK_CODUSU) <> "ADMIN" And Trim(LK_CODUSU) <> "SUPERVISOR" Then
 Exit Sub
End If
If Trim(lblnom(Index).Tag) = "" Then
 Exit Sub
End If
Dim wnombre
wnombre = InputBox("Ingrese la Nueva Descripción para este Campo :", Pub_Titulo, Trim(lblnom(Index).Caption))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
Screen.MousePointer = 11
SQ_OPER = 1
PUB_TIPREG = LOC_TIPREG
PUB_NUMTAB = Val(lblnom(Index).Tag)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_llave.EOF Then
  tab_llave.AddNew
Else
  tab_llave.Edit
End If
  tab_llave!tab_codcia = LK_CODCIA
  tab_llave!tab_tipreg = LOC_TIPREG
  tab_llave!TAB_NUMTAB = Val(lblnom(Index).Tag)
  tab_llave!TAB_NOMLARGO = Left(wnombre, 40)
  tab_llave!TAB_NOMCORTO = Left(wnombre, 10)
  tab_llave.Update
  lblnom(Index).Caption = Left(wnombre, 40)
Screen.MousePointer = 0
End Sub

Private Sub letraotorgado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.letraotorgado.TabIndex
End If

End Sub

Private Sub ListExiste_Click()
Dim d, C, a As Integer
End Sub
Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmCLI.F14.Visible = False
     frmCLI.txtesposo.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub ListExiste_LostFocus()
If frmCLI.ListExiste.Visible = False Then
    Exit Sub
End If
End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 Txt_key.text = Trim(ListView1.ListItems.Item(loc_key).text) & " "
 txt_key_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.text = Trim(ListView1.ListItems.Item(loc_key).text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 Txt_key.text = ""
 Txt_key.SetFocus
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick

End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub

Private Sub otrocontrato_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.otrocontrato.TabIndex
End If
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True
 If CU > 4 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If
End Sub

Public Sub ASIGNA_INT(wcontrol As ComboBox, txt As Integer)
For fila = 0 To wcontrol.ListCount - 1
    If Val(Trim(Right(wcontrol.List(fila), 3))) = txt Then
        wcontrol.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub
Public Sub ASIGNA_subgrupo(wcontrol As ComboBox, txt As String)
For fila = 0 To wcontrol.ListCount - 1
    If Val(Trim(Right(wcontrol.List(fila), 3))) = Val(txt) Then
        wcontrol.ListIndex = fila
        Exit Sub
    End If
Next fila
End Sub

Public Sub LLENA_CLI(ban As Integer, CG As String)
    If ban = 0 Then
        '**  BAN = 0 BUSCA DATOS NUEVAMENTE
        If loc_key > ListView1.ListItems.Count Or loc_key = 0 Then
         Else
          Txt_key.text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
        End If
        pu_cp = Left(CmbCGP.text, 2)
        pu_codclie = Val(Txt_key.text)
        SQ_OPER = 1
        pu_codcia = LK_CODCIA
        LEER_CLILOC_LLAVE
    End If
    loc_ultcod = Val(cliloc_llave!CLI_CODCLIE)
    frmCLI.Txt_key.text = cliloc_llave!CLI_CODCLIE
    LLENA_123
    txtnombre.text = Nulo_Valors(cliloc_llave!cli_nombre)
    txtesposo.text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_ESPOSO))
    Txtesposa.text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_ESPOSA))
    TxtEmpresa.text = Trim(Nulo_Valors(cliloc_llave!CLI_NOMBRE_EMPRESA))
    Txtdireccion.text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA_DIREC))
    Txtnumdir.text = Trim(Nulo_Valor0(cliloc_llave!CLI_CASA_NUM))
    ASIGNA_INT TxtZona, Nulo_Valor0(cliloc_llave!CLI_CASA_ZONA)
    ASIGNA_INT TxtSubZona, Nulo_Valor0(cliloc_llave!CLI_CASA_SUBZONA)
    ASIGNA_INT txtZonaNew, Nulo_Valor0(cliloc_llave!CLI_ZONA_NEW)
    txtDirTrabajo.text = Trim(Nulo_Valors(cliloc_llave!CLI_TRAB_DIREC))
    txtnumdirtrabajo.text = Trim(Nulo_Valor0(cliloc_llave!CLI_TRAB_NUM))
    ASIGNA_INT TxtZonaTrabajo, Nulo_Valor0(cliloc_llave!cli_TRAB_ZONA)
    ASIGNA_INT TxtSubZonaTrabajo, Nulo_Valor0(cliloc_llave!cli_TRAB_SUBZONA)
    ASIGNA_INT TxtLugarCasa, Nulo_Valor0(cliloc_llave!CLI_LUGAR_CASA)
    ASIGNA_INT TxtLugarTrab, Nulo_Valor0(cliloc_llave!CLI_LUGAR_TRAB)
       
    txtRUCesposo.text = Trim(Nulo_Valors(cliloc_llave!cli_RUC_ESPOSO))
    txtRUCesposa.text = Trim(Nulo_Valors(cliloc_llave!CLI_RUC_ESPOSA))
    txtRUCempresa.text = Trim(Nulo_Valors(cliloc_llave!CLI_RUC_EMPRESA))
    frmCLI.txtpropiedad1.text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA1))
    frmCLI.txtpropiedad2.text = Trim(Nulo_Valors(cliloc_llave!CLI_CASA2))
    frmCLI.txtregpublico1.text = Trim(Nulo_Valors(cliloc_llave!CLI_REGPUB1))
    frmCLI.txtregpublico2.text = Trim(Nulo_Valors(cliloc_llave!CLI_REGPUB2))
    frmCLI.txtautovaluo.text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTOAVALUO))
    frmCLI.txtauto1.text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTO1))
    frmCLI.txtauto2.text = Trim(Nulo_Valors(cliloc_llave!CLI_AUTO2))
    frmCLI.txtprendas.text = Trim(Nulo_Valors(cliloc_llave!CLI_PRENDA))
    frmCLI.txttelefono1.text = Trim(Nulo_Valors(cliloc_llave!CLI_TELEF1))
    frmCLI.txttelefono2.text = Trim(Nulo_Valors(cliloc_llave!CLI_TELEF2))
    frmCLI.otrocontrato.Value = Nulo_Valor0(cliloc_llave!CLI_OTRO_CONTR)
    frmCLI.letraotorgado.Value = Nulo_Valor0(cliloc_llave!CLI_LETRA)
    LLENA_BLOQ
    ASIGNA_INT cmbgrupo, Nulo_Valors(cliloc_llave!CLI_GRUPO)
    ASIGNA_subgrupo txtsubgrupo, Nulo_Valors(cliloc_llave!CLI_SUBGRUPO)
    frmCLI.txtNucleo.text = Nulo_Valor0(cliloc_llave!CLI_nucleo)
    frmCLI.txtestado.text = Nulo_Valors(cliloc_llave!CLI_estado)
    frmCLI.txtprog.text = Nulo_Valors(cliloc_llave!CLI_programado)
    frmCLI.tcuenta.text = Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB)
    If Trim(Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB)) <> "" Then
        cmdcontab.Caption = "&Quitar Relacion Contable"
    Else
        cmdcontab.Caption = "Relacionar a Con&tabilidad"
    End If
    frmCLI.tcuenta2.text = Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB2)
    If Trim(Nulo_Valors(cliloc_llave!CLI_CUENTA_CONTAB2)) <> "" Then
        cmdcontab2.Caption = "&Quitar Relacion Contable"
    Else
        cmdcontab2.Caption = "Relacionar a Con&tabilidad"
    End If
    frmCLI.txtlimite.text = Nulo_Valor0(cliloc_llave!cli_limcre)
    If Nulo_Valors(cliloc_llave!cli_moneda) = "S" Then
       cli_moneda.ListIndex = 0
    ElseIf cliloc_llave!cli_moneda = "D" Then
       cli_moneda.ListIndex = 1
    End If
    txtDTX.text = Nulo_Valors(cliloc_llave!CLI_DET_TOT)
    
End Sub

Public Sub LIMPIA_CLI()
    Txt_key.text = ""
    txtnombre.text = ""
    txtesposo.text = ""
    Txtesposa.text = ""
    TxtEmpresa.text = ""
    Txtdireccion.text = ""
    Txtnumdir.text = ""
    TxtZona.ListIndex = -1
    TxtSubZona.ListIndex = -1
    txtZonaNew.ListIndex = -1
    TxtLugarCasa.ListIndex = -1
    TxtLugarTrab.ListIndex = -1
    txtDirTrabajo.text = ""
    txtnumdirtrabajo.text = ""
    frmCLI.TxtZonaTrabajo.ListIndex = -1
    TxtSubZonaTrabajo.ListIndex = -1
    txtRUCesposo.text = ""
    txtRUCesposa.text = ""
    txtRUCempresa.text = ""
    frmCLI.txtpropiedad2.text = ""
    frmCLI.txtpropiedad1.text = ""
    frmCLI.txtregpublico1.text = ""
    frmCLI.txtregpublico2.text = ""
    frmCLI.txtautovaluo.text = ""
    frmCLI.txtauto1.text = ""
    frmCLI.txtauto2.text = ""
    frmCLI.txtprendas.text = ""
    frmCLI.txttelefono1.text = ""
    frmCLI.txttelefono2.text = ""
    frmCLI.otrocontrato.Value = 0
    frmCLI.letraotorgado.Value = 0
    frmCLI.ListBloqueos.Clear
    frmCLI.cmbgrupo.ListIndex = -1
    frmCLI.txtsubgrupo.ListIndex = -1
    frmCLI.txtNucleo.text = ""
    frmCLI.txtestado.text = ""
    frmCLI.tcuenta.text = ""
    frmCLI.OptNombre(0).Value = False
    frmCLI.OptNombre(1).Value = False
    frmCLI.OptNombre(2).Value = False
    frmCLI.txtlimite.text = ""
    cli_moneda.ListIndex = -1
    frmCLI.txtDTX = ""
    txtprog.text = ""
    LOC_CTA_CLI = ""
    LOC_DES_CLI = ""
    tcuenta2.text = ""
    LOC_CTA_CLI2 = ""
    LOC_DES_CLI2 = ""
    cmdcontab.Caption = "Relacionar a Con&tabilidad"
    cmdcontab2.Caption = "Relacionar a Con&tabilidad"

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
 If txtesposo.Enabled And txtesposo.Visible Then
   txtesposo.SetFocus
 End If
Else
 If txtDirTrabajo.Enabled And txtDirTrabajo.Visible Then
   txtDirTrabajo.SetFocus
 End If
End If
End Sub

Private Sub SSTab1_GotFocus()
If ListView1.Visible Then
 frmCLI.Txt_key.text = ""
 frmCLI.ListView1.Visible = False
End If
End Sub

Private Sub tcuenta_GotFocus()
Azul tcuenta, tcuenta
End Sub

Private Sub tcuenta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If frmCLI.cmdModificar.Enabled Then
          frmCLI.cmdModificar.SetFocus
    Else
           frmCLI.CmdAgregar.SetFocus
    End If
 End If
End Sub

Private Sub txt_key_GotFocus()
 Azul Txt_key, Txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.

If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo FIN
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.Count Then loc_key = ListView1.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.Count Then loc_key = ListView1.ListItems.Count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo FIN
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  Txt_key.text = Trim(ListView1.ListItems.Item(loc_key).text) & " "
  Txt_key.SelStart = Len(Txt_key.text)
FIN:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
'On Error GoTo SALCODI
If KeyAscii = 27 And Trim(txtnombre.text) = "" Then
 Txt_key.text = ""
End If
If KeyAscii <> 13 Or Left(CmdAgregar.Caption, 2) = "&G" Or Left(cmdModificar.Caption, 2) = "&G" Then
   GoTo FIN
End If
pu_codclie = Val(Txt_key.text)
If Len(Txt_key.text) = 0 Then
   Exit Sub
End If
fra2.Refresh
If pu_codclie <> 0 And IsNumeric(Txt_key.text) = True Then
   SQ_OPER = 1
   On Error GoTo mucho
   pu_codcia = LK_CODCIA
   pu_cp = Left(CmbCGP.text, 1)
   pu_codclie = Val(Txt_key.text)
   LEER_CLILOC_LLAVE
   On Error GoTo 0
   If cliloc_llave.EOF Then
     Azul Txt_key, Txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Txt_key.SetFocus
     GoTo FIN
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   If Left(CmbCGP.text, 1) = "C" Then
         LLENA_CLI 1, "C"
   End If
   If Left(CmbCGP.text, 1) = "P" Then
         LLENA_CLI 1, "P"
   End If
   frmCLI.Txt_key.Locked = True
   frmCLI.cmdModificar.SetFocus
   Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.Count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).text)
   If Trim(UCase(Txt_key.text)) = Left(valor, Len(Trim(Txt_key.text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdCancelar.Enabled = True
   If Left(CmbCGP.text, 1) = "C" Then
         LLENA_CLI 0, "C"
   End If
   If Left(CmbCGP.text, 1) = "P" Then
         LLENA_CLI 0, "P"
   End If
   frmCLI.Txt_key.Locked = True
   cmdCancelar.Enabled = True
   frmCLI.cmdModificar.SetFocus
End If
dale:
ListView1.Visible = False
FIN:
mucho:
Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmCLI
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim VAR
If Len(Txt_key.text) = 0 Or IsNumeric(Txt_key.text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.text) = 1 Then
    If Txt_key.text = "" Then Txt_key.text = " "
    VAR = Asc(Txt_key.text)
    VAR = VAR + 1
    NADA = VAR
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM FROM CLIENTES WHERE  CLI_CP = '" & Left(CmbCGP.text, 1) & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & Txt_key.text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        Txt_key.text = ""
      End If
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(Txt_key.text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.Count Then
      ListView1.ListItems.Item(ListView1.ListItems.Count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If
End Sub

Private Sub txtauto1_GotFocus()
Azul txtauto1, txtauto1
End Sub

Private Sub txtauto1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtauto1.TabIndex
End If
End Sub

Private Sub txtauto2_GotFocus()
Azul txtauto2, txtauto2
End Sub

Private Sub txtauto2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SIGUE_CAMPO frmCLI.txtauto2.TabIndex
End If
End Sub

Private Sub txtautovaluo_GotFocus()
Azul txtautovaluo, txtautovaluo
End Sub

Private Sub txtautovaluo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SIGUE_CAMPO frmCLI.txtautovaluo.TabIndex
End If
End Sub
Private Sub Txtdireccion_GotFocus()
Azul Txtdireccion, Txtdireccion
fra1.Refresh
End Sub

Private Sub txtdireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.TxtLugarCasa.SetFocus
    SendKeys "%{up}"
End If
End Sub

Private Sub Txtdireccion_LostFocus()
If Left(CmbCGP.text, 1) = "C" Then
  If Trim(txtDirTrabajo.text) = "" Then
    txtDirTrabajo.text = Trim(Txtdireccion.text)
  End If
End If
End Sub

Private Sub txtDirTrabajo_GotFocus()
Azul txtDirTrabajo, txtDirTrabajo
fra2.Refresh
End Sub

Private Sub txtDirTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtDirTrabajo.TabIndex
End If
End Sub

Private Sub txtDTX_GotFocus()
Azul txtDTX, txtDTX
End Sub

Private Sub txtDTX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtDTX.TabIndex
End If

End Sub

Private Sub TxtEmpresa_GotFocus()
 Azul TxtEmpresa, TxtEmpresa
End Sub

Private Sub TxtEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.txtRUCempresa.SetFocus
End If
End Sub

Private Sub TxtEmpresa_LostFocus()
fra1.Refresh
End Sub

Private Sub Txtesposa_GotFocus()
 Azul Txtesposa, Txtesposa
End Sub

Private Sub Txtesposa_KeyPress(KeyAscii As Integer)
  
If KeyAscii = 13 Then
    If frmCLI.txtRUCesposa.Visible Then
       frmCLI.txtRUCesposa.SetFocus
    Else
        frmCLI.txttelefono2.SetFocus
    End If
End If
End Sub

Private Sub Txtesposa_LostFocus()
fra1.Refresh
End Sub

Private Sub txtesposo_GotFocus()
 Azul txtesposo, txtesposo
End Sub

Private Sub txtesposo_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
  KeyAscii = 0
  Exit Sub
End If
If KeyAscii = 13 Then
  frmCLI.txtRUCesposo.SetFocus
End If
End Sub

Public Sub GRABAR_CLI(wCGH As String)
Dim VAR_CIAS As String * 2
Dim TOTCIAS As String
Dim xcuenta As Integer
Dim Modo As String * 1
PS_GEN(0) = 0
gen.Requery
TOTCIAS = Nulo_Valors(gen!gen_cli_cias)
xcuenta = 1
For fila = 1 To 30
    If Trim(Mid(TOTCIAS, xcuenta, 2)) = LK_CODCIA Then
         GoTo SIGUE_PASA
    End If
    xcuenta = xcuenta + 2
Next fila
GoTo CIA_ACTUAL

SIGUE_PASA:
If Trim(TOTCIAS) <> "" And Left(CmbCGP.text, 1) = "C" Then
  xcuenta = 1
  For fila = 1 To 30
    If Trim(Mid(TOTCIAS, xcuenta, 2)) = "" Then
      Exit For
    Else
       PSPAR_CLI(0) = Mid(TOTCIAS, xcuenta, 2)
       par_llave_cli.Requery
       If par_llave_cli.EOF Then
       '     MsgBox "No Grabo en la Compañia : " + Mid(TOTCIAS, xcuenta, 2) + " No Existe", 48, Pub_Titulo
       Else
           VAR_CIAS = Mid(TOTCIAS, xcuenta, 2)
           If Left(cmdModificar.Caption, 2) = "&G" Then
             If VAR_CIAS = LK_CODCIA Then GoTo PASA
             SQ_OPER = 1
             pu_cp = wCGH
             pu_codclie = Val(frmCLI.Txt_key.text)
             pu_codcia = VAR_CIAS
             LEER_CLILOC_LLAVE
             If cliloc_llave.EOF Then
'                MsgBox "No Grabo en la Compañia : " + VAR_CIAS + " No Existe cliente ", 48, Pub_Titulo
             Else
               cliloc_llave.Edit
               Modo = "E"
               GoSub Grabar
               cliloc_llave.Update
             End If
PASA:
           Else
             cliloc_llave.AddNew
             Modo = "A"
             GoSub Grabar
             cliloc_llave.Update
           End If
      End If
    End If
    xcuenta = xcuenta + 2
  Next fila
  'ACTUALIZA POR ULTIMO LA CIA ACTUAL PARA MANTENER LA LLAVE ACTIVA
  If Left(cmdModificar.Caption, 2) = "&G" Then
    VAR_CIAS = LK_CODCIA
    SQ_OPER = 1
    pu_cp = wCGH
    pu_codclie = Val(frmCLI.Txt_key.text)
    pu_codcia = VAR_CIAS
    LEER_CLILOC_LLAVE
    If cliloc_llave.EOF Then
      MsgBox "No Grabo en la Compañia : " + VAR_CIAS + " No Existe cliente ", 48, Pub_Titulo
    Else
      cliloc_llave.Edit
      cliloc_llave!cli_limcre = Val(frmCLI.txtlimite.text)
      Modo = "E"
      GoSub Grabar
      cliloc_llave.Update
    End If
  End If
  
Else
CIA_ACTUAL:
  VAR_CIAS = LK_CODCIA
  If Left(cmdModificar.Caption, 2) = "&G" Then
    cliloc_llave.Edit
    Modo = "E"
  Else
    GRABA_CONTAB VAR_CIAS
    cliloc_llave.AddNew
    Modo = "A"
  End If
  GoSub Grabar
  cliloc_llave.Update

  Exit Sub
End If
 
Exit Sub
   
Grabar:
    If Modo = "A" Then
       cliloc_llave!CLI_CP = wCGH
       cliloc_llave!CLI_CODCLIE = Val(frmCLI.Txt_key.text)
       cliloc_llave!cli_SALDO = 0
       cliloc_llave!CLI_DET_TOT = "D"
       cliloc_llave!cli_moneda = "S"
       If Left(CmbCGP.text, 1) = "C" Then
        loc_ultcod = Val(frmCLI.Txt_key.text)
       End If
    End If
    cliloc_llave!CLI_CODCIA = VAR_CIAS
    cliloc_llave!CLI_NOMBRE_ESPOSO = txtesposo.text
    cliloc_llave!CLI_NOMBRE_ESPOSA = Txtesposa.text
    cliloc_llave!CLI_NOMBRE_EMPRESA = TxtEmpresa.text
    ASIGNA_123
    cliloc_llave!cli_nombre = frmCLI.txtnombre.text
    cliloc_llave!CLI_CASA_DIREC = Txtdireccion.text
    cliloc_llave!CLI_CASA_NUM = Val(Txtnumdir.text)
    cliloc_llave!CLI_CASA_ZONA = Val(Right(TxtZona.text, 4))
    cliloc_llave!CLI_LUGAR_CASA = Val(Right(TxtLugarCasa.text, 4))
    cliloc_llave!CLI_LUGAR_TRAB = Val(Right(TxtLugarTrab.text, 4))
    cliloc_llave!CLI_CASA_SUBZONA = Val(Right(TxtSubZona.text, 4))
    cliloc_llave!CLI_ZONA_NEW = Val(Right(txtZonaNew.text, 4))
    cliloc_llave!CLI_TRAB_DIREC = txtDirTrabajo.text
    cliloc_llave!CLI_TRAB_NUM = Nulo_Valor0(txtnumdirtrabajo.text)
    cliloc_llave!cli_TRAB_ZONA = Val(Right(frmCLI.TxtZonaTrabajo.text, 4))
    cliloc_llave!cli_TRAB_SUBZONA = Val(Right(TxtSubZonaTrabajo.text, 4))
    cliloc_llave!cli_RUC_ESPOSO = txtRUCesposo.text
    cliloc_llave!CLI_RUC_ESPOSA = txtRUCesposa.text
    cliloc_llave!CLI_RUC_EMPRESA = txtRUCempresa.text
    cliloc_llave!CLI_CASA1 = frmCLI.txtpropiedad1.text
    cliloc_llave!CLI_CASA2 = frmCLI.txtpropiedad2.text
    cliloc_llave!CLI_REGPUB1 = frmCLI.txtregpublico1.text
    cliloc_llave!CLI_REGPUB2 = frmCLI.txtregpublico2.text
    cliloc_llave!CLI_AUTOAVALUO = frmCLI.txtautovaluo.text
    cliloc_llave!CLI_AUTO1 = frmCLI.txtauto1.text
    cliloc_llave!CLI_AUTO2 = frmCLI.txtauto2.text
    cliloc_llave!CLI_PRENDA = frmCLI.txtprendas.text
    cliloc_llave!CLI_TELEF1 = frmCLI.txttelefono1.text
    cliloc_llave!CLI_TELEF2 = frmCLI.txttelefono2.text
    cliloc_llave!CLI_OTRO_CONTR = frmCLI.otrocontrato.Value
    cliloc_llave!CLI_LETRA = frmCLI.letraotorgado.Value
    cliloc_llave!CLI_GRUPO = Val(Right(frmCLI.cmbgrupo.text, 4))
    cliloc_llave!CLI_SUBGRUPO = Val(Right(frmCLI.txtsubgrupo.text, 4))
    cliloc_llave!CLI_CORRELA = 0
    cliloc_llave!CLI_nucleo = frmCLI.txtNucleo.text
    cliloc_llave!CLI_estado = frmCLI.txtestado.text
    cliloc_llave!CLI_programado = Nulo_Valors(txtprog.text)
    '  <<< Actualiza La Cta. solo de la Cia Actual >>>
    If VAR_CIAS = LK_CODCIA Then
      cliloc_llave!CLI_CUENTA_CONTAB = Trim(frmCLI.tcuenta.text)
      cliloc_llave!CLI_CUENTA_CONTAB2 = Trim(frmCLI.tcuenta2.text)
    End If
    cliloc_llave!cli_moneda = Left(cli_moneda.text, 1)
    If txtDTX.text = "" Then
     txtDTX.text = " "
    End If
    cliloc_llave!CLI_DET_TOT = txtDTX.text
    If Trim(TOTCIAS) = "" Then
      cliloc_llave!cli_limcre = Val(txtlimite.text)
    End If
Return
End Sub

Public Sub MENSAJE_CLI(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub
Public Function GENERA_PRO() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim i, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double
pu_cp = "P"
pu_codclie = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
LEER_CLILOC_LLAVE
If cliloc_mayor.EOF Then
    NUMCAD = "1"
Else
    cliloc_mayor.MoveLast
    NUMCAD = cliloc_mayor!CLI_CODCLIE
End If
VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If

For i = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, i, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next i
VINT3 = VINT2 * 13 - 5

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)
GENERA_PRO = INTpub_cadena
End Function

Public Function GENERA_CODI() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim i, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double
pu_cp = "C"
pu_codclie = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
LEER_CLILOC_LLAVE

If cliloc_mayor.EOF Then
    NUMCAD = "1"
Else
    cliloc_mayor.MoveLast
    NUMCAD = cliloc_mayor!CLI_CODCLIE
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If
For i = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, i, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next i
VINT3 = VINT2 * 9

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)

GENERA_CODI = INTpub_cadena

End Function
Public Function CONSIS_CLI() As Boolean
If frmCLI.OptNombre(0).Value Then
    If Trim(frmCLI.txtesposo.text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        txtesposo.SetFocus
        GoTo ESCAPA
    End If
ElseIf frmCLI.OptNombre(1).Value Then
    If Trim(frmCLI.Txtesposa.text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        Txtesposa.SetFocus
        GoTo ESCAPA
    End If
ElseIf frmCLI.OptNombre(2).Value Then
    If Trim(frmCLI.TxtEmpresa.text) = "" Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Datos Principal ..."
        TxtEmpresa.SetFocus
        GoTo ESCAPA
    End If
End If

If Len(frmCLI.txtesposo.text) = 0 And Len(frmCLI.Txtesposa.text) = 0 And Len(frmCLI.TxtEmpresa.text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Ingrese Algun Nombre  ..."
        txtesposo.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(0).Value And Len(frmCLI.txtesposo.text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        txtesposo.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(1).Value And Len(frmCLI.Txtesposa.text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        Txtesposa.SetFocus
        GoTo ESCAPA
ElseIf frmCLI.OptNombre(2).Value And Len(frmCLI.TxtEmpresa.text) = 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "Nombre  NO Puede estar en Blanco ..."
        TxtEmpresa.SetFocus
        GoTo ESCAPA
ElseIf Len(frmCLI.txtesposo.text) = 0 And Len(frmCLI.txtRUCesposo.text) > 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "RUC  debe estar en Blanco ..."
        txtRUCesposo.SetFocus
        GoTo ESCAPA
ElseIf Len(frmCLI.Txtesposa.text) = 0 And Len(frmCLI.txtRUCesposa.text) > 0 Then
        CONSIS_CLI = False
        MENSAJE_CLI "L.E.  debe estar en Blanco ..."
        txtRUCesposa.SetFocus
        GoTo ESCAPA
ElseIf Len(frmCLI.TxtEmpresa.text) = 0 And Len(frmCLI.txtRUCempresa.text) > 0 Then
     If LK_EMP <> "PLA" Then
        CONSIS_CLI = False
        MENSAJE_CLI "RUC  debe estar en Blanco ..."
        txtRUCempresa.SetFocus
        GoTo ESCAPA
     End If
End If

If frmCLI.txtRUCesposo.text <> "" Then
    If Len(Trim(frmCLI.txtRUCesposo.text)) <> 8 Then
       CONSIS_CLI = False
       MENSAJE_CLI "R.U.C. de No es Validad ..."
       frmCLI.txtRUCesposo.SetFocus
       GoTo ESCAPA
    End If
End If
If frmCLI.txtRUCesposa.text <> "" Then
    If Len(Trim(frmCLI.txtRUCesposa.text)) <> 8 Then
       CONSIS_CLI = False
       MENSAJE_CLI "L.E. de No es Validad ..."
       frmCLI.txtRUCesposa.SetFocus
       GoTo ESCAPA
    End If
End If
If LK_EMP <> "PLA" Then
 If Left(CmbCGP.text, 1) = "C" Then
  If frmCLI.txtRUCempresa.text <> "" Then
    If Len(Trim(frmCLI.txtRUCempresa.text)) <> 8 Then
       CONSIS_CLI = False
       MENSAJE_CLI "L.E. de No es Validad ..."
       txtRUCempresa.SetFocus
       GoTo ESCAPA
    End If
  End If
 End If
End If
If LK_EMP = "HER" And frmCLI.txtRUCesposo.text <> "" Then
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = frmCLI.txtRUCesposo.text
 PS_REP01(2) = frmCLI.Txt_key.text
 llave_rep01.Requery
 If Not llave_rep01.EOF Then
   MsgBox "RUC Existe en otro Cliente : " + Trim(llave_rep01!cli_nombre), 48, Pub_Titulo
    CONSIS_CLI = False
    Azul frmCLI.txtRUCesposo, frmCLI.txtRUCesposo
    GoTo ESCAPA
 End If
End If
If Trim(tcuenta.text) <> "" Then
' SQ_OPER = 1
' PUB_CUENTA = Trim(tcuenta.text)
' LEER_COM_LLAVE
' If com_llave.EOF Then
'  MsgBox "Cuanta Contable No Existe. Verificar ", 48, Pub_Titulo
'  CONSIS_CLI = False
'  Azul frmCLI.tcuenta, frmCLI.tcuenta
'  GoTo ESCAPA
 'End If
End If
If frmCLI.txtNucleo = "" Then
   frmCLI.txtNucleo.text = " "
End If
If frmCLI.txtestado = "" Then
   frmCLI.txtestado.text = "A"
End If

If Left(CmbCGP.text, 1) = "C" Then
 If Trim(TxtLugarCasa.text) = "" Then
    MsgBox "Dato no es opcional ,Lugar.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtLugarCasa.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtZona.text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtZona.SetFocus
    GoTo ESCAPA
 End If
 If Trim(TxtSubZona.text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    TxtSubZona.SetFocus
    GoTo ESCAPA
 End If
 If Trim(txtZonaNew.text) = "" Then
    MsgBox "Dato no es opcional ,Definir.", 48, Pub_Titulo
    CONSIS_CLI = False
    txtZonaNew.SetFocus
    GoTo ESCAPA
 End If
End If

CONSIS_CLI = True
ESCAPA:
End Function

Private Sub txtesposo_LostFocus()
fra1.Refresh
End Sub

Private Sub txtestado_GotFocus()
Azul txtestado, txtestado
End Sub

Private Sub txtestado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtestado.TabIndex
   fra2.Refresh
End If
End Sub

Private Sub txtlimite_GotFocus()
Azul txtlimite, txtlimite
End Sub

Private Sub txtlimite_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Left(CmdAgregar.Caption, 2) = "&G" Then
   CmdAgregar.SetFocus
 Else
   cmdModificar.SetFocus
 End If
End If

End Sub

Private Sub TxtLugarCasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.Txtnumdir.SetFocus
End If

End Sub

Private Sub TxtLugarCasa_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtLugarCasa.ListIndex
PUB_TIPREG = Mid(TxtLugarCasa.ToolTipText, 13, Len(TxtLugarCasa.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtLugarCasa, 25
LLENA_ZONA TxtLugarTrab, 25
TxtLugarCasa.SetFocus
SendKeys "%{up}"

End Sub

Private Sub TxtLugarCasa_LostFocus()
On Error GoTo sigue
If Left(CmbCGP.text, 1) = "C" Then
  If Trim(TxtLugarTrab.text) = "" Then
    TxtLugarTrab.ListIndex = TxtLugarCasa.ListIndex
  End If
End If
fra1.Refresh
Exit Sub
sigue:
End Sub

Private Sub TxtLugarTrab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.TxtLugarTrab.TabIndex
   fra2.Refresh
End If

End Sub

Private Sub TxtLugarTrab_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtLugarTrab.ListIndex
PUB_TIPREG = Mid(TxtLugarTrab.ToolTipText, 13, Len(TxtLugarTrab.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtLugarCasa, 25
LLENA_ZONA TxtLugarTrab, 25
TxtLugarTrab.SetFocus
SendKeys "%{up}"

End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtesposo.SetFocus
  fra2.Refresh
End If
End Sub

Private Sub txtNucleo_GotFocus()
Azul txtNucleo, txtNucleo
End Sub

Private Sub txtnucleo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtNucleo.TabIndex
   fra2.Refresh
End If

End Sub

Private Sub Txtnumdir_GotFocus()
Azul Txtnumdir, Txtnumdir
fra1.Refresh
End Sub

Private Sub Txtnumdir_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
   txttelefono1.SetFocus
End If
End Sub

Private Sub Txtnumdir_LostFocus()
On Error GoTo sigue
If Left(CmbCGP.text, 1) = "C" Then
  If Val(txtnumdirtrabajo.text) = 0 Then
    txtnumdirtrabajo.text = Txtnumdir.text
  End If
End If
Exit Sub
sigue:
End Sub

Private Sub txtnumdirtrabajo_GotFocus()
Azul txtnumdirtrabajo, txtnumdirtrabajo
fra2.Refresh
End Sub

Private Sub txtnumdirtrabajo_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  SIGUE_CAMPO frmCLI.txtnumdirtrabajo.TabIndex
End If
End Sub

Private Sub txtprendas_GotFocus()
Azul txtprendas, txtprendas
End Sub

Private Sub txtprendas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SIGUE_CAMPO frmCLI.txtprendas.TabIndex
End If
End Sub

Private Sub txtprog_GotFocus()
 Azul txtprog, txtprog
End Sub

Private Sub txtprog_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtprog.TabIndex
   fra2.Refresh
End If

End Sub

Private Sub txtpropiedad1_GotFocus()
Azul txtpropiedad1, txtpropiedad1
End Sub

Private Sub txtpropiedad1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SIGUE_CAMPO frmCLI.txtpropiedad1.TabIndex
  fra2.Refresh
End If
End Sub

Private Sub txtpropiedad2_GotFocus()
Azul txtpropiedad2, txtpropiedad2
End Sub

Private Sub txtpropiedad2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtpropiedad2.TabIndex
End If
End Sub

Private Sub txtregpublico1_GotFocus()
Azul txtregpublico1, txtregpublico1
End Sub

Private Sub txtregpublico1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtregpublico1.TabIndex
   fra2.Refresh
End If
End Sub

Private Sub txtregpublico2_GotFocus()
Azul txtregpublico2, txtregpublico2
End Sub

Private Sub txtregpublico2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.txtregpublico2.TabIndex
End If
End Sub

Private Sub txtRUCempresa_GotFocus()
Azul txtRUCempresa, txtRUCempresa
End Sub

Private Sub txtRUCempresa_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
    frmCLI.Txtdireccion.SetFocus
End If
End Sub

Private Sub txtRUCesposa_GotFocus()
Azul txtRUCesposa, txtRUCesposa
End Sub

Private Sub txtRUCesposa_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
    frmCLI.TxtEmpresa.SetFocus
End If
End Sub

Private Sub txtRUCesposo_GotFocus()
 Azul txtRUCesposo, txtRUCesposo
End Sub

Private Sub txtRUCesposo_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
   frmCLI.Txtesposa.SetFocus
End If
End Sub

Private Sub txtsubgrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
' frmCLI.SSTab1.Tab = 1
If txtlimite.Enabled And txtlimite.Visible Then
 txtlimite.SetFocus
End If
End If
End Sub

Private Sub txtsubgrupo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = txtsubgrupo.ListIndex
PUB_TIPREG = Mid(txtsubgrupo.ToolTipText, 13, Len(txtsubgrupo.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "SUB - GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_GRUPOS txtsubgrupo, 333
txtsubgrupo.SetFocus
SendKeys "%{up}"
End Sub

Private Sub TxtSubZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.txtZonaNew.SetFocus
    SendKeys "%{UP}"
End If
End Sub

Private Sub TxtSubZona_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtSubZona.ListIndex
PUB_TIPREG = Mid(TxtSubZona.ToolTipText, 13, Len(TxtSubZona.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtSubZona, 30
LLENA_ZONA TxtSubZonaTrabajo, 35
TxtSubZona.SetFocus
SendKeys "%{up}"


End Sub

Private Sub TxtSubZona_LostFocus()
fra1.Refresh
End Sub

Private Sub TxtSubZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SIGUE_CAMPO TxtSubZonaTrabajo.TabIndex
End If
End Sub

Private Sub TxtSubZonaTrabajo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtSubZonaTrabajo.ListIndex
PUB_TIPREG = Mid(TxtSubZonaTrabajo.ToolTipText, 13, Len(TxtSubZonaTrabajo.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtSubZonaTrabajo, 35
LLENA_ZONA TxtSubZona, 30

TxtSubZonaTrabajo.SetFocus
SendKeys "%{up}"

End Sub

Private Sub txttelefono1_GotFocus()
Azul txttelefono1, txttelefono1
End Sub

Private Sub txttelefono1_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  frmCLI.TxtZona.SetFocus
  SendKeys "%{up}"
End If
End Sub

Private Sub txttelefono1_LostFocus()
On Error GoTo sigue
If Left(CmbCGP.text, 1) = "C" Then
  If Trim(txttelefono2.text) = "" Then
    txttelefono2.text = txttelefono1.text
  End If
End If
fra1.Refresh
Exit Sub
sigue:
End Sub

Private Sub txttelefono2_GotFocus()
Azul txttelefono2, txttelefono2
fra2.Refresh
End Sub

Private Sub txttelefono2_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  SIGUE_CAMPO frmCLI.txttelefono2.TabIndex
End If
End Sub

Private Sub TxtZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.TxtSubZona.SetFocus
    SendKeys "%{up}"
End If
End Sub

Private Sub TxtZona_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtZona.ListIndex
PUB_TIPREG = Mid(TxtZona.ToolTipText, 13, Len(TxtZona.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtZona, 20
LLENA_ZONA TxtZonaTrabajo, 20
TxtZona.SetFocus
SendKeys "%{up}"

End Sub

Private Sub TxtZona_LostFocus()
On Error GoTo sigue
If Left(CmbCGP.text, 1) = "C" Then
  If Trim(TxtZonaTrabajo.text) = "" Then
      TxtZonaTrabajo.ListIndex = TxtZona.ListIndex
  End If
End If
fra1.Refresh
Exit Sub
sigue:


fra1.Refresh
End Sub

Private Sub txtZonaNew_LostFocus()
On Error GoTo sigue
If Left(CmbCGP.text, 1) = "C" Then
  If Trim(TxtSubZonaTrabajo.text) = "" Then
      TxtSubZonaTrabajo.ListIndex = txtZonaNew.ListIndex
  End If
End If
fra1.Refresh
Exit Sub
sigue:
End Sub

Private Sub TxtZonaTrabajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SIGUE_CAMPO frmCLI.TxtZonaTrabajo.TabIndex
End If
End Sub

Public Function EXISTE_CLI(wCP As String, VALOR1 As String, WCODI As String) As Boolean
Dim VAR
Dim tempo
tempo = Left(Trim(VALOR1), Len(VALOR1) - 1)
VAR = Asc(Right(Trim(VALOR1), 1))
VAR = VAR + 1
If VAR = 91 Then
  VAR = "ZZZZZZZZ"
Else
  VAR = Chr(VAR)
End If
tempo = tempo + VAR
archi = "SELECT * FROM CLIENTES WHERE  CLI_CODCLIE <> " & WCODI & " AND CLI_CP = '" & wCP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & VALOR1 & "' AND  '" & tempo & "' ORDER BY CLI_NOMBRE"
ListExiste.Clear
Set PSX = CN.CreateQuery("", archi)
Set X = PSX.OpenResultset(rdOpenKeyset)
X.Requery
EXISTE_CLI = False
If X.EOF Then
 frmCLI.ListExiste.Clear
 GoTo FIN
End If

If wCP = "P" Then
    F14.Caption = "Lista de Proveedores parecidos ... "
End If
frmCLI.ListExiste.TextMatrix(0, 0) = "Cia"
frmCLI.ListExiste.TextMatrix(0, 1) = "Codigo "
If wCP = "C" Then
    F14.Caption = "Lista de Clientes parecidos ... "
End If
frmCLI.ListExiste.TextMatrix(0, 2) = lblnom(0).Caption
frmCLI.ListExiste.TextMatrix(0, 3) = lblnom(2).Caption
frmCLI.ListExiste.TextMatrix(0, 4) = lblnom(6).Caption & " " & lblnom(7).Caption

fila = 0
frmCLI.ListExiste.Rows = 2
Do Until X.EOF
    fila = fila + 1
    frmCLI.ListExiste.TextMatrix(fila, 0) = Nulo_Valors(X!CLI_CODCIA)
    frmCLI.ListExiste.TextMatrix(fila, 1) = Nulo_Valors(X!CLI_CODCLIE)
    frmCLI.ListExiste.TextMatrix(fila, 2) = Nulo_Valors(X!CLI_NOMBRE_ESPOSO)
    frmCLI.ListExiste.TextMatrix(fila, 3) = Nulo_Valors(X!CLI_NOMBRE_ESPOSA)
    frmCLI.ListExiste.TextMatrix(fila, 4) = Nulo_Valors(X!CLI_CASA_DIREC) & "  # " & Nulo_Valors(X!CLI_CASA_NUM)
    EXISTE_CLI = True
    frmCLI.ListExiste.Rows = frmCLI.ListExiste.Rows + 1
    X.MoveNext
Loop

If EXISTE_CLI Then
    frmCLI.ListExiste.Rows = frmCLI.ListExiste.Rows - 1
    Op(0).Value = False
    Op(0).Enabled = False
    Op(1).Value = True
    frmCLI.F14.Visible = True
    frmCLI.ListExiste.Row = 1
    frmCLI.ListExiste.Col = 1
    frmCLI.ListExiste.SetFocus
End If
GoTo FIN
Exit Function

CHECKERROR:
MsgBox Err.Description
FIN:

End Function

Private Sub TxtZonaTrabajo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
wpos = TxtZonaTrabajo.ListIndex
PUB_TIPREG = Mid(TxtZonaTrabajo.ToolTipText, 13, Len(TxtZonaTrabajo.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA TxtZonaTrabajo, 20
LLENA_ZONA TxtZona, 20
TxtZonaTrabajo.SetFocus
SendKeys "%{up}"
End Sub
Private Sub TxtZonanew_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmCLI.cmbgrupo.SetFocus
    SendKeys "%{UP}"
End If
End Sub

Private Sub TxtZonanew_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos As Integer
If KeyCode <> 45 Then
  Exit Sub
End If
PUB_TIPREG = Mid(txtZonaNew.ToolTipText, 13, Len(txtZonaNew.ToolTipText))
PUB_CODCIA = "00"
Load FrmDatArti
FrmDatArti.Caption = "ZONAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENA_ZONA txtZonaNew, 35
txtZonaNew.SetFocus
SendKeys "%{up}"
fra2.Refresh
End Sub


Public Sub ETIQUETA_CLI()
SQ_OPER = 1
PUB_TIPREG = LOC_TIPREG
PUB_CODCIA = LK_CODCIA
For fila = 0 To lblnom.Count - 1
 PUB_NUMTAB = Val(lblnom(fila).Tag)
 LEER_TAB_LLAVE
 If tab_llave.EOF Then
 Else
 If fila = 30 Then
' Stop
 End If
  lblnom(fila).Caption = Trim(tab_llave!TAB_NOMLARGO)
 End If
Next fila
End Sub

Public Sub LEER_CLILOC_LLAVE()
Select Case SQ_OPER
Case 1
PSCLILOC_LLAVE.rdoParameters(0) = pu_cp
PSCLILOC_LLAVE.rdoParameters(1) = pu_codclie
PSCLILOC_LLAVE.rdoParameters(2) = pu_codcia
cliloc_llave.Requery
GoTo salida

Case 2
PSCLILOC_MAYOR.rdoParameters(0) = pu_cp
PSCLILOC_MAYOR.rdoParameters(1) = pu_codclie
PSCLILOC_MAYOR.rdoParameters(2) = pu_codcia
cliloc_mayor.Requery
GoTo salida

Case 3
GoTo salida

End Select


salida:

End Sub

Public Sub ACCESO_CLI()
On Error GoTo Ver
Dim tAcceso As String
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
tAcceso = Trim(Nulo_Valors(par_llave!par_acceso_cli))

WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(tAcceso)
cade = Trim(tAcceso)
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
   ' WNUM tiene el codigo a mostrar
   GoSub muestracli
Loop

Exit Sub

muestracli:
fila = 0
Do Until fila >= frmCLI.Controls.Count
 If TypeOf frmCLI.Controls(fila) Is Timer Then
   GoTo OTRITO
 End If
 If TypeOf frmCLI.Controls(fila) Is MSFlexGrid Then
   GoTo OTRITO
 End If
 If TypeOf frmCLI.Controls(fila) Is Line Then
   GoTo OTRITO
 End If
 If frmCLI.Controls(fila).WhatsThisHelpID = WNUM Then
   frmCLI.Controls(fila).Visible = True
 End If
 
OTRITO:
 fila = fila + 1
Loop
Return
Exit Sub
Ver:
MsgBox Err.Description, 48, Pub_Titulo
Resume Next

End Sub

Public Sub SIGUE_CAMPO(WTAG As Integer)
' 40 ES EL MAXIMO DE CAMPOS DISPONIBLE
Dim wmax As Integer
Dim cuenta As Integer
wmax = 42
fila = WTAG
Do Until fila >= wmax
 fila = fila + 1
 cuenta = 0
 Do Until cuenta >= frmCLI.Controls.Count - 1
  If TypeOf frmCLI.Controls(cuenta) Is Timer Then
    GoTo OTRITO
  End If
  If TypeOf frmCLI.Controls(cuenta) Is MSFlexGrid Then
    GoTo OTRITO
  End If
'  MsgBox frmCLI.Controls(fila).Name
  If TypeOf frmCLI.Controls(cuenta) Is OptionButton Then
    GoTo OTRITO
  End If
  If frmCLI.Controls(cuenta).TabIndex = fila Then
    If frmCLI.Controls(cuenta).Visible Then
         frmCLI.Controls(cuenta).SetFocus
         Exit Sub
    End If
  End If
OTRITO:
  cuenta = cuenta + 1
 Loop
Loop
If frmCLI.cmdModificar.Enabled Then
   frmCLI.cmdModificar.SetFocus
Else
   frmCLI.CmdAgregar.SetFocus
End If
End Sub

Public Sub GRABA_CONTAB(wcia As String)
 If Left(CmbCGP.text, 1) = "C" Then
   If Nulo_Valors(par_llave!PAR_CONTA_C) <> "A" Then
     Exit Sub
   End If
 ElseIf Left(CmbCGP.text, 1) = "P" Then
   If Nulo_Valors(par_llave!PAR_CONTA_P) <> "A" Then
     Exit Sub
   End If
 End If
If Trim(LOC_CTA_CLI) <> "" Then
 If Trim(tcuenta.text) <> "" Then
   On Error GoTo cuenta1
  com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI
    com_llave!com_DESCRIPCION = LOC_DES_CLI
    com_llave!COM_NIVEL = LOC_NIVEL
    com_llave!com_cuenta_sup = LOC_CTA_SUP
    com_llave!com_FLAG_AFECTACION = LOC_FLAG_AFEC
    com_llave!com_ESTADO = LOC_ESTADO
    com_llave!COM_TIPO_CTA = LOC_TIPO_CTA
    com_llave!com_SIGNO_D = LOC_SIGNO_D
    com_llave!com_SIGNO_H = LOC_SIGNO_H
    com_llave!com_ACT_PAS = LOC_ACT_PAS
    com_llave!com_SIGNO_H = LOC_SIGNO_H
    com_llave!com_ACT_PAS = LOC_ACT_PAS
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = ""
    com_llave!com_cuenta_AUTO_H = ""
  com_llave.Update
  On Error GoTo 0
  cmdcontab.Caption = "&Quitar Relación Contable"
 End If
End If

Exit Sub
If Trim(LOC_CTA_CLI2) <> "" Then
 If Trim(tcuenta2.text) <> "" Then
    On Error GoTo CUENTA2
    com_llave.AddNew
    com_llave!COM_CODCIA = wcia
    com_llave!com_cuenta = LOC_CTA_CLI2
    com_llave!com_DESCRIPCION = LOC_DES_CLI2
    com_llave!COM_NIVEL = LOC_NIVEL2
    com_llave!com_cuenta_sup = LOC_CTA_SUP2
    com_llave!com_FLAG_AFECTACION = LOC_FLAG_AFEC2
    com_llave!com_ESTADO = LOC_ESTADO2
    com_llave!COM_TIPO_CTA = LOC_TIPO_CTA2
    com_llave!com_SIGNO_D = LOC_SIGNO_D2
    com_llave!com_SIGNO_H = LOC_SIGNO_H2
    com_llave!com_ACT_PAS = LOC_ACT_PAS2
    com_llave!com_SIGNO_H = LOC_SIGNO_H2
    com_llave!com_ACT_PAS = LOC_ACT_PAS2
    com_llave!COM_DEB_MES = 0
    com_llave!COM_HAB_MES = 0
    com_llave!COM_HAB_ANO = 0
    com_llave!COM_DEB_ANO = 0
    com_llave!com_cuenta_AUTOM_D = ""
    com_llave!com_cuenta_AUTO_H = ""
    com_llave.Update
    On Error GoTo 0
    cmdcontab2.Caption = "&Quitar Relación Contable"
 End If
End If
Exit Sub

cuenta1:
If Err.Number = 40002 Then
  MsgBox "Cuenta Existe, NO Procede. Cta.: " & LOC_CTA_CLI, 48, Pub_Titulo
   tcuenta.text = ""
  com_llave.CancelUpdate
End If
Exit Sub

CUENTA2:
If Err.Number = 40002 Then
  MsgBox "Cuenta Existe, NO Procede. Cta.: " & LOC_CTA_CLI2, 48, Pub_Titulo
  tcuenta2.text = ""
  com_llave.CancelUpdate
End If
Exit Sub

End Sub
